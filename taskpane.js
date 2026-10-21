"use strict";
const HELPER = "http://localhost:8081";
const SKIP_PARAS = 40; // so viele Absätze hinter der Stelle (bei Bedarf anpassen: größer = weiterer Hop)
let seq = -1;
let timer = null;
let lastErr = "";
const byId = (id) => document.getElementById(id);
const out = (m) => { byId("log").textContent += m + "\n"; };
Office.onReady(() => {
    try {   // Add-in beim Öffnen des Dokuments automatisch laden (nur mit Shared Runtime; sonst wirkungslos)
        const sb = Office.addin && Office.addin.setStartupBehavior;
        if (sb) { const r = sb.call(Office.addin, Office.StartupBehavior.load); if (r && r.catch) r.catch(() => { }); }
    }
    catch (e) { }
    byId("btn-sets").onclick = () => run(checkSets);
    byId("btn-ping").onclick = () => run(ping);
    byId("btn-jump").onclick = () => run(() => jump(byId("term").value, (parseInt(byId("occ").value) || 1) - 1));
    byId("btn-now").onclick = () => run(() => poll(true));
    byId("chk-follow").onchange = (e) => setFollow(e.target.checked);
    setFollow(true); // automatisch starten: kein Klick mehr nötig
});
function setFollow(on) {
    if (timer) {
        clearInterval(timer);
        timer = null;
    }
    byId("chk-follow").checked = on;
    if (on) {
        seq = -1;
        timer = setInterval(() => run(() => poll(false)), 800);
        out("Folgen: an");
        setStatus("…verbinde");
    }
    else {
        out("Folgen: aus");
        setStatus("Folgen aus");
    }
}
function setStatus(m) { byId("status").textContent = m; }
async function run(fn) {
    try {
        await fn();
        lastErr = "";
    }
    catch (e) {
        const m = String(e);
        setStatus("○ Helfer nicht erreichbar");
        if (m !== lastErr) {
            out("FEHLER: " + m);
            lastErr = m;
        }
    }
}
function checkSets() {
    const sets = [["WordApi", "1.1"], ["WordApi", "1.3"], ["WordApi", "1.4"], ["WordApi", "1.5"], ["WordApiDesktop", "1.1"], ["WordApiHiddenDocument", "1.4"]];
    out("Plattform: " + Office.context.platform);
    sets.forEach(s => out(s[0] + " " + s[1] + ": " + Office.context.requirements.isSetSupported(s[0], s[1])));
}
async function ping() {
    const r = await fetch(HELPER + "/ping", { targetAddressSpace: "loopback" });
    out("Helfer erreichbar: " + (await r.text()));
}
// Suchmuster, das alle Bindestrich-Varianten und feste Leerzeichen toleriert (Word-Platzhaltersuche)
function fuzzy(term) {
    return term.split("").map((ch) => {
        if (/[-\u2010-\u2015\u2212]/.test(ch))
            return "[\u2010\u2011\u2012\u2013\u2014\u2212-]";
        if (/[\s\u00a0]/.test(ch))
            return "[ \u00a0]";
        if ("?*[]<>(){}@!^\\".indexOf(ch) >= 0)
            return "\\" + ch;
        return ch;
    }).join("");
}
async function jump(term, occ) {
    await Word.run(async (context) => {
        const body = context.document.body;
        let res = body.search(term.replace(/\^/g, "^^").replace(/\u00a0/g, "^s"), { matchCase: false });
        res.load("items");
        await context.sync();
        let how = "exakt";
        if (!res.items.length) {
            res = body.search(fuzzy(term), { matchCase: false, matchWildcards: true });
            res.load("items");
            await context.sync();
            how = "tolerant";
        }
        out("Suche \"" + term + "\" (" + how + "): " + res.items.length + " Treffer");
        if (!res.items.length)
            return;
        const i = Math.min(occ, res.items.length - 1);
        if (byId("chk-top").checked) {
            // Word scrollt beim Markieren nur "gerade eben sichtbar" (Stelle landet unten).
            // Trick: kurz zu einem Absatz etwa eine Bildschirmhöhe HINTER der Stelle springen,
            // dann zur Stelle zurück -> sie landet am oberen Rand (kleiner Hop statt Dokumentende).
            let p = res.items[i].paragraphs.getFirst();
            const ahead = [];
            for (let k = 0; k < SKIP_PARAS; k++) {
                p = p.getNextOrNullObject();
                ahead.push(p);
                p.load("isNullObject");
            }
            await context.sync();
            let far = null;
            for (let k = ahead.length - 1; k >= 0; k--) {
                if (!ahead[k].isNullObject) {
                    far = ahead[k];
                    break;
                }
            }
            if (far) {
                far.getRange("Start").select();
                await context.sync();
            }
        }
        res.items[i].select();
        await context.sync();
        out("  -> Treffer " + (i + 1) + " ausgewählt");
    });
}
async function poll(force) {
    const r = await fetch(HELPER + "/target", { targetAddressSpace: "loopback" });
    const t = await r.json();
    setStatus("● folgt dem PageTurner" + (t.term ? " – " + t.term : ""));
    if (seq === -1 && !force) {
        seq = t.seq;
        return;
    } // beim Start nicht zu einem alten Ziel springen
    if (force || t.seq !== seq) {
        seq = t.seq;
        if (t.term) {
            out("Ziel " + seq + " (" + t.src + "): " + t.term + " #" + (t.occ + 1));
            await jump(t.term, t.occ);
        }
        else if (force) {
            out("Noch kein Ziel gemeldet");
        }
    }
}
