import http from 'node:http';
import { WebSocketServer, WebSocket } from 'ws';

const PORT = process.env.PORT || 8080;
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const MODEL = process.env.MODEL || 'gpt-realtime';
const VOICE = process.env.VOICE || 'alloy';
const PROMPT_URL =
  process.env.PROMPT_URL ||
  'https://appriver3651000954-my.sharepoint.com/personal/jacks_palittoconsulting_com/_layouts/15/download.aspx?share=IQDp5lIcOXxlRqnNoaQTkUXJAXgAZ261M3yeO35W63F9gew';

const DEFAULT_PROMPT =
  "You are a friendly phone assistant. As soon as the call connects, greet the caller first by saying: \"Thank you for calling. How can I help?\" Then assist them.";

if (!OPENAI_API_KEY) {
  console.error('FATAL: OPENAI_API_KEY env var is not set.');
  process.exit(1);
}

// --- live prompt fetch (so you can edit SharePoint without redeploying) ---
async function fetchPrompt() {
  try {
    const res = await fetch(PROMPT_URL, { redirect: 'follow' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const text = (await res.text()).trim();
    if (!text || text.length < 5) throw new Error('empty body');
    console.log('Prompt fetched, length', text.length);
    return text;
  } catch (e) {
    console.error('Prompt fetch failed, using default:', e.message);
    return DEFAULT_PROMPT;
  }
}

// --- TwiML served to Twilio: connect the call's media to this server ---
const twiml = (host) =>
  `<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Connect>
    <Stream url="wss://${host}/twilio" />
  </Connect>
</Response>`;

const server = http.createServer((req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);

  if (url.pathname === '/voice' && (req.method === 'POST' || req.method === 'GET')) {
    res.writeHead(200, { 'Content-Type': 'text/xml' });
    res.end(twiml(req.headers.host));
    return;
  }
  if (url.pathname === '/health') {
    res.writeHead(200, { 'Content-Type': 'text/plain' });
    res.end('ok');
    return;
  }
  res.writeHead(404);
  res.end('not found');
});

const wss = new WebSocketServer({ server, path: '/twilio' });

wss.on('connection', (twilioWs) => {
  console.log('── Twilio WS connected ──');

  let streamSid = null;
  let openaiWs = null;
  let sessionReady = false; // true after session.updated (audio formats locked in)
  let greeted = false;
  const audioQueue = []; // caller audio captured before OpenAI is ready

  // 1) Open the OpenAI Realtime socket for this call
  (async () => {
    const instructions = await fetchPrompt();

    openaiWs = new WebSocket(`wss://api.openai.com/v1/realtime?model=${MODEL}`, {
      headers: {
        Authorization: `Bearer ${OPENAI_API_KEY}`
      },
    });

    openaiWs.on('open', () => console.log('OpenAI WS open'));

    openaiWs.on('message', (raw) => {
      let msg;
      try {
        msg = JSON.parse(raw.toString());
      } catch {
        return;
      }

      switch (msg.type) {
        // Configure the session: g711 µ-law both ways (pure passthrough with Twilio)
        case 'session.created': {
          openaiWs.send(
            JSON.stringify({
              type: 'session.update',
              session: {
                type: 'realtime',
                instructions,
                audio: {
                  input: {
                    format: { type: 'audio/pcmu' }, // µ-law 8kHz
                    turn_detection: {
                      type: 'server_vad',
                      threshold: 0.5,
                      prefix_padding_ms: 300,
                      silence_duration_ms: 500,
                    },
                  },
                  output: {
                    format: { type: 'audio/pcmu' }, // µ-law 8kHz
                    voice: VOICE,
                  },
                },
              },
            })
          );
          break;
        }

        // Session is fully configured → safe to send audio and trigger the greeting
        case 'session.updated': {
          if (!sessionReady) {
            sessionReady = true;
            console.log('Session ready');

            // flush any caller audio we buffered while waiting
            for (const a of audioQueue) {
              openaiWs.send(JSON.stringify({ type: 'input_audio_buffer.append', audio: a }));
            }
            audioQueue.length = 0;

            // GREETING-FIRST: trigger the model to speak. No hardcoded text —
            // it generates the opening from your SharePoint instructions.
            if (!greeted) {
              greeted = true;
              openaiWs.send(JSON.stringify({ type: 'response.create' }));
            }
          }
          break;
        }

        // OpenAI speech → relay straight to Twilio (already µ-law base64)
        case 'response.output_audio.delta': {
          if (streamSid && msg.delta) {
            twilioWs.send(
              JSON.stringify({
                event: 'media',
                streamSid,
                media: { payload: msg.delta },
              })
            );
          }
          break;
        }

        // Caller started talking → barge-in: dump whatever Twilio still has queued
        case 'input_audio_buffer.speech_started': {
          if (streamSid) {
            twilioWs.send(JSON.stringify({ event: 'clear', streamSid }));
          }
          break;
        }

        case 'error': {
          console.error('OpenAI error:', JSON.stringify(msg.error || msg));
          break;
        }
      }
    });

    openaiWs.on('close', (code, reason) => {
      console.log('OpenAI WS closed', code, reason?.toString());
      try {
        twilioWs.close();
      } catch {}
    });
    openaiWs.on('error', (e) => console.error('OpenAI WS error:', e.message));
  })();

  // 2) Handle Twilio media-stream frames
  twilioWs.on('message', (raw) => {
    let msg;
    try {
      msg = JSON.parse(raw.toString());
    } catch {
      return;
    }

    switch (msg.event) {
      case 'connected':
        console.log('Twilio: connected');
        break;

      case 'start':
        streamSid = msg.start.streamSid;
        console.log('Twilio: start, streamSid =', streamSid);
        break;

      case 'media': {
        const payload = msg.media?.payload; // base64 µ-law 8kHz
        if (!payload) break;
        if (openaiWs && openaiWs.readyState === WebSocket.OPEN && sessionReady) {
          openaiWs.send(JSON.stringify({ type: 'input_audio_buffer.append', audio: payload }));
        } else {
          if (audioQueue.length < 1000) audioQueue.push(payload); // safety cap
        }
        break;
      }

      case 'stop':
        console.log('Twilio: stop');
        try {
          openaiWs && openaiWs.close();
        } catch {}
        break;
    }
  });

  twilioWs.on('close', () => {
    console.log('── Twilio WS closed ──');
    try {
      openaiWs && openaiWs.close();
    } catch {}
  });
  twilioWs.on('error', (e) => console.error('Twilio WS error:', e.message));
});

server.listen(PORT, () => console.log('Relay listening on', PORT));