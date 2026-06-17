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

const DEPARTMENTS = [
  'Customer Service',
  'Insurance',
  'Branch Representative',
  'Mortgage Payments',
];

// Safety: if Twilio never echoes the mark, transfer anyway after this long.
const TRANSFER_MARK_TIMEOUT_MS = 8000;

if (!OPENAI_API_KEY) {
  console.error('FATAL: OPENAI_API_KEY env var is not set.');
  process.exit(1);
}

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

  const host = req.headers.host;
  let streamSid = null;
  let callSid = null;
  let openaiWs = null;
  let sessionReady = false;
  let greeted = false;
  let transferring = false;

  // pending transfer state
  let pendingDepartment = null; // set when model calls transfer_call
  let markSent = false;
  let transferTimer = null;

  const audioQueue = [];

  function doTransfer() {
    if (transferring) return;
    transferring = true;
    if (transferTimer) clearTimeout(transferTimer);
    console.log('Executing transfer →', pendingDepartment);
    transferCall(callSid, host, pendingDepartment);
  }

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
                      'Transfer the caller to a department. First tell the caller you are transferring them, then call this function. The system will play the audio fully before transferring.',
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
          // Don't barge-in/clear if we're mid-transfer-wait.
          if (!transferring && !pendingDepartment && streamSid) {
            twilioWs.send(JSON.stringify({ event: 'clear', streamSid }));
          }
          break;
        }

        // Catch the function call (set pending; do NOT transfer yet)
        case 'response.output_item.done': {
          const item = msg.item;
          if (
            item &&
            item.type === 'function_call' &&
            item.name === 'transfer_call'
          ) {
            let department = 'the requested department';
            try {
              const args = JSON.parse(item.arguments || '{}');
              if (args.department) department = args.department;
            } catch {}
            if (!pendingDepartment) {
              pendingDepartment = department;
              console.log('Transfer pending →', department, '(waiting for audio to finish)');
            }
          }
          break;
        }

        // Response fully generated. If a transfer is pending, mark the audio
        // stream so we know when Twilio has PLAYED everything.
        case 'response.done': {
          if (pendingDepartment && !markSent && !transferring) {
            markSent = true;
            if (streamSid) {
              twilioWs.send(
                JSON.stringify({
                  event: 'mark',
                  streamSid,
                  mark: { name: 'transfer_ready' },
                })
              );
              console.log('Mark sent, waiting for Twilio playback to finish…');
              // Safety net if the mark never comes back.
              transferTimer = setTimeout(() => {
                console.log('Mark timeout — transferring anyway.');
                doTransfer();
              }, TRANSFER_MARK_TIMEOUT_MS);
            } else {
              doTransfer();
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
        callSid = msg.start.callSid;
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

      // Twilio echoes our mark once it has PLAYED all audio up to that point.
      case 'mark': {
        if (msg.mark?.name === 'transfer_ready' && pendingDepartment) {
          console.log('Twilio playback finished (mark received) → transferring.');
          doTransfer();
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
    if (transferTimer) clearTimeout(transferTimer);
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