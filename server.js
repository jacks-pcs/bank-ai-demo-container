import http from 'node:http';
import { WebSocketServer, WebSocket } from 'ws';

const PORT = process.env.PORT || 8080;
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
const TWILIO_ACCOUNT_SID = process.env.TWILIO_ACCOUNT_SID;
const TWILIO_AUTH_TOKEN = process.env.TWILIO_AUTH_TOKEN;
const MODEL = process.env.MODEL || 'gpt-realtime';
const VOICE = process.env.VOICE || 'alloy';
const PROMPT_URL =
  process.env.PROMPT_URL ||
  'https://appriver3651000954-my.sharepoint.com/personal/jacks_palittoconsulting_com/_layouts/15/download.aspx?share=IQDp5lIcOXxlRqnNoaQTkUXJAXgAZ261M3yeO35W63F9gew';

const DEFAULT_PROMPT =
  "You are a friendly phone assistant. As soon as the call connects, greet the caller first by saying: \"Thank you for calling. How can I help?\" Then assist them.";

// The departments the AI can transfer to.
const DEPARTMENTS = [
  'Customer Service',
  'Insurance',
  'Branch Representative',
  'Mortgage Payments',
];

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

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// --- Twilio REST: redirect the live call to new TwiML (the "transfer") ---
async function transferCall(callSid, host, department) {
  if (!TWILIO_ACCOUNT_SID || !TWILIO_AUTH_TOKEN) {
    console.error('Cannot transfer: Twilio creds not set.');
    return;
  }
  if (!callSid) {
    console.error('Cannot transfer: no callSid.');
    return;
  }

  const dept = xmlEscape(department);
  const twiml = `<?xml version="1.0" encoding="UTF-8"?>
<Response>
  <Say>You have been transferred to ${dept}. Redirecting back to the main menu.</Say>
  <Redirect>https://${host}/voice</Redirect>
</Response>`;

  const auth = Buffer.from(`${TWILIO_ACCOUNT_SID}:${TWILIO_AUTH_TOKEN}`).toString('base64');
  const body = new URLSearchParams({ Twiml: twiml });

  try {
    const res = await fetch(
      `https://api.twilio.com/2010-04-01/Accounts/${TWILIO_ACCOUNT_SID}/Calls/${callSid}.json`,
      {
        method: 'POST',
        headers: {
          Authorization: `Basic ${auth}`,
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body,
      }
    );
    if (!res.ok) {
      console.error('Twilio redirect failed:', res.status, await res.text());
    } else {
      console.log('Transfer redirect sent for', department);
    }
  } catch (e) {
    console.error('Twilio redirect error:', e.message);
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

wss.on('connection', (twilioWs, req) => {
  console.log('── Twilio WS connected ──');

  const host = req.headers.host; // used to build the /voice redirect URL
  let streamSid = null;
  let callSid = null;
  let openaiWs = null;
  let sessionReady = false;
  let greeted = false;
  let transferring = false;
  const audioQueue = [];

  // 1) Open the OpenAI Realtime socket for this call
  (async () => {
    const instructions = await fetchPrompt();

    openaiWs = new WebSocket(`wss://api.openai.com/v1/realtime?model=${MODEL}`, {
      headers: {
        Authorization: `Bearer ${OPENAI_API_KEY}`,
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
        case 'session.created': {
          openaiWs.send(
            JSON.stringify({
              type: 'session.update',
              session: {
                type: 'realtime',
                instructions,
                tools: [
                  {
                    type: 'function',
                    name: 'transfer_call',
                    description:
                      'Transfer the caller to a department. Call this as soon as the caller asks to be transferred, or when their request clearly belongs to one of the departments.',
                    parameters: {
                      type: 'object',
                      properties: {
                        department: {
                          type: 'string',
                          enum: DEPARTMENTS,
                          description: 'Which department to transfer the caller to.',
                        },
                      },
                      required: ['department'],
                    },
                  },
                ],
                tool_choice: 'auto',
                audio: {
                  input: {
                    format: { type: 'audio/pcmu' },
                    turn_detection: {
                      type: 'server_vad',
                      threshold: 0.5,
                      prefix_padding_ms: 300,
                      silence_duration_ms: 500,
                    },
                  },
                  output: {
                    format: { type: 'audio/pcmu' },
                    voice: VOICE,
                  },
                },
              },
            })
          );
          break;
        }

        case 'session.updated': {
          if (!sessionReady) {
            sessionReady = true;
            console.log('Session ready');

            for (const a of audioQueue) {
              openaiWs.send(
                JSON.stringify({ type: 'input_audio_buffer.append', audio: a })
              );
            }
            audioQueue.length = 0;

            if (!greeted) {
              greeted = true;
              openaiWs.send(JSON.stringify({ type: 'response.create' }));
            }
          }
          break;
        }

        case 'response.output_audio.delta': {
          if (!transferring && streamSid && msg.delta) {
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

        case 'input_audio_buffer.speech_started': {
          if (!transferring && streamSid) {
            twilioWs.send(JSON.stringify({ event: 'clear', streamSid }));
          }
          break;
        }

        // The model called a function → handle transfer
        case 'response.output_item.done': {
          const item = msg.item;
          if (item && item.type === 'function_call' && item.name === 'transfer_call') {
            let department = 'the requested department';
            try {
              const args = JSON.parse(item.arguments || '{}');
              if (args.department) department = args.department;
            } catch {}

            if (!transferring) {
              transferring = true;
              console.log('Transfer requested →', department);
              // Redirect the live Twilio call. This tears down the stream + session,
              // speaks the line, then /voice restarts a fresh realtime session.
              transferCall(callSid, host, department);
            }
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
        callSid = msg.start.callSid; // needed for the REST transfer
        console.log('Twilio: start, streamSid =', streamSid, 'callSid =', callSid);
        break;

      case 'media': {
        const payload = msg.media?.payload;
        if (!payload) break;
        if (openaiWs && openaiWs.readyState === WebSocket.OPEN && sessionReady) {
          openaiWs.send(
            JSON.stringify({ type: 'input_audio_buffer.append', audio: payload })
          );
        } else {
          if (audioQueue.length < 1000) audioQueue.push(payload);
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

server.on('upgrade', (req) => {
  console.log('HTTP upgrade request for:', req.url);
});

server.listen(PORT, () => console.log('Relay listening on', PORT, '— build OK'));