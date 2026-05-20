// heartbeat.js — sends POST /api/addin/heartbeat to the proxy on a 15-second interval.
// Calls onResult({ seq, ok, serverTime?, rtt?, error? }) after each attempt.

import { CONFIG } from './config.js';

const INTERVAL_MS = 15_000;

let _seq   = 0;
let _timer = null;

async function beat(onResult) {
    const seq    = ++_seq;
    const sentAt = Date.now();
    try {
        const res = await fetch(`${CONFIG.PROXY_URL}/api/addin/heartbeat`, {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify({ seq, clientTime: new Date().toISOString() }),
            signal:  AbortSignal.timeout(10_000),
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        const rtt  = Date.now() - sentAt;
        console.log(`[Heartbeat #${seq}] ack=${data.ack} serverTime=${data.serverTime} rtt=${rtt}ms`);
        onResult({ seq, ok: true, serverTime: data.serverTime, rtt });
    } catch (e) {
        console.warn(`[Heartbeat #${seq}] failed: ${e.message}`);
        onResult({ seq, ok: false, error: e.message });
    }
}

export function startHeartbeat(onResult) {
    beat(onResult); // fire immediately on load
    _timer = setInterval(() => beat(onResult), INTERVAL_MS);
}

export function stopHeartbeat() {
    if (_timer) { clearInterval(_timer); _timer = null; }
}
