/* global Word, Office */
"use strict";

// ── Office initialisieren ────────────────────────────────────────────────────

Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
        document.getElementById("scan-btn").addEventListener("click", scanDocument);
        document.getElementById("fill-btn").addEventListener("click", fillPlaceholders);
        scanDocument();   // automatisch beim Öffnen einlesen
    }
});

// ── Platzhalter im Dokument suchen ───────────────────────────────────────────

async function scanDocument() {
    showStatus("Dokument wird durchsucht …", "info");
    document.getElementById("fill-btn").classList.add("hidden");
    document.getElementById("fields-container").innerHTML = "";

    try {
        var found = new Set();

        await Word.run(async function (context) {

            // 1) Fliesstext (Body)
            var body = context.document.body;
            body.load("text");
            await context.sync();
            collectPlaceholders(body.text, found);

            // 2) Textfelder / Shapes  (Word 2019+ / M365)
            try {
                var shapes = body.shapes;
                shapes.load("items/textFrame/textRange/text");
                await context.sync();
                for (var i = 0; i < shapes.items.length; i++) {
                    var shp = shapes.items[i];
                    if (shp.textFrame && shp.textFrame.textRange) {
                        collectPlaceholders(shp.textFrame.textRange.text, found);
                    }
                }
            } catch (_e) {
                // Shapes-API nicht verfügbar – wird still ignoriert
            }

            // 3) Kopf- und Fusszeilen
            try {
                var headers = context.document.sections.getFirst().getHeader("primary");
                var footers = context.document.sections.getFirst().getFooter("primary");
                headers.load("text");
                footers.load("text");
                await context.sync();
                collectPlaceholders(headers.text, found);
                collectPlaceholders(footers.text, found);
            } catch (_e2) {
                // ignorieren
            }
        });

        buildForm(Array.from(found));

    } catch (err) {
        showStatus("Fehler beim Lesen des Dokuments: " + err.message, "error");
    }
}

function collectPlaceholders(text, setObj) {
    var regex = /\[[^\]]+\]/g;
    var match;
    while ((match = regex.exec(text)) !== null) {
        setObj.add(match[0]);
    }
}

// ── Eingabeformular aufbauen ─────────────────────────────────────────────────

function buildForm(placeholders) {
    var container = document.getElementById("fields-container");
    var fillBtn   = document.getElementById("fill-btn");
    container.innerHTML = "";

    if (placeholders.length === 0) {
        container.innerHTML = '<p class="no-placeholders">Keine Platzhalter&nbsp;[&nbsp;] gefunden.</p>';
        fillBtn.classList.add("hidden");
        showStatus("", "");
        return;
    }

    // Trennlinie mit Anzahl
    var heading = document.createElement("p");
    heading.className = "section-title";
    heading.textContent = placeholders.length + " Platzhalter gefunden";
    container.appendChild(heading);

    placeholders.forEach(function (ph, i) {
        var inner = ph.slice(1, -1);   // [ ] entfernen

        var row = document.createElement("div");
        row.className = "field-row";

        var lbl = document.createElement("label");
        lbl.htmlFor     = "ctrl_" + i;
        lbl.textContent = inner;
        row.appendChild(lbl);

        var ctrl;
        if (inner.indexOf("/") !== -1) {
            // Dropdown – Optionen durch "/" getrennt
            ctrl = document.createElement("select");
            inner.split("/").forEach(function (opt) {
                var option = document.createElement("option");
                option.value       = opt.trim();
                option.textContent = opt.trim();
                ctrl.appendChild(option);
            });
        } else {
            // Freitexteingabe
            ctrl = document.createElement("input");
            ctrl.type        = "text";
            ctrl.placeholder = inner + " …";
        }

        ctrl.id                  = "ctrl_" + i;
        ctrl.dataset.placeholder = ph;
        row.appendChild(ctrl);
        container.appendChild(row);
    });

    fillBtn.classList.remove("hidden");
    showStatus("", "");
}

// ── Platzhalter im Dokument ersetzen ─────────────────────────────────────────

async function fillPlaceholders() {
    var inputs = document.querySelectorAll("#fields-container [data-placeholder]");
    if (!inputs.length) return;

    var fillBtn = document.getElementById("fill-btn");
    fillBtn.disabled    = true;
    fillBtn.textContent = "Wird ersetzt …";

    var totalReplaced = 0;
    var shapeWarning  = false;

    try {
        await Word.run(async function (context) {

            for (var idx = 0; idx < inputs.length; idx++) {
                var input = inputs[idx];
                var ph    = input.dataset.placeholder;
                var value = (input.tagName === "SELECT")
                    ? input.options[input.selectedIndex].value
                    : input.value.trim();

                if (!value) continue;

                // ── Body-Text ──
                var results = context.document.body.search(ph, { matchCase: true });
                results.load("items");
                await context.sync();

                results.items.forEach(function (r) {
                    r.insertText(value, Word.InsertLocation.replace);
                });
                totalReplaced += results.items.length;
                await context.sync();
            }

            // ── Textfelder / Shapes ──
            try {
                var shapes = context.document.body.shapes;
                shapes.load("items/textFrame/textRange");
                await context.sync();

                for (var si = 0; si < shapes.items.length; si++) {
                    var shp = shapes.items[si];
                    if (!shp.textFrame || !shp.textFrame.textRange) continue;

                    for (var idx2 = 0; idx2 < inputs.length; idx2++) {
                        var input2 = inputs[idx2];
                        var ph2    = input2.dataset.placeholder;
                        var value2 = (input2.tagName === "SELECT")
                            ? input2.options[input2.selectedIndex].value
                            : input2.value.trim();

                        if (!value2) continue;

                        var sr = shp.textFrame.textRange.search(ph2, { matchCase: true });
                        sr.load("items");
                        await context.sync();
                        sr.items.forEach(function (r) {
                            r.insertText(value2, Word.InsertLocation.replace);
                        });
                        totalReplaced += sr.items.length;
                        await context.sync();
                    }
                }
            } catch (_shapeErr) {
                shapeWarning = true;   // Shapes-API nicht verfügbar
            }

            // ── Kopf-/Fusszeilen ──
            try {
                for (var idx3 = 0; idx3 < inputs.length; idx3++) {
                    var input3 = inputs[idx3];
                    var ph3    = input3.dataset.placeholder;
                    var value3 = (input3.tagName === "SELECT")
                        ? input3.options[input3.selectedIndex].value
                        : input3.value.trim();

                    if (!value3) continue;

                    var section = context.document.sections.getFirst();
                    var hdr = section.getHeader("primary");
                    var ftr = section.getFooter("primary");

                    var hr = hdr.search(ph3, { matchCase: true });
                    var fr = ftr.search(ph3, { matchCase: true });
                    hr.load("items");
                    fr.load("items");
                    await context.sync();

                    hr.items.forEach(function (r) { r.insertText(value3, Word.InsertLocation.replace); });
                    fr.items.forEach(function (r) { r.insertText(value3, Word.InsertLocation.replace); });
                    totalReplaced += hr.items.length + fr.items.length;
                    await context.sync();
                }
            } catch (_hdrErr) {
                // ignorieren
            }
        });

        // ── Ergebnis anzeigen ──
        var msg = totalReplaced + " Ersetzung" + (totalReplaced !== 1 ? "en" : "") + " vorgenommen.";
        if (shapeWarning) {
            msg += " Hinweis: Textfelder bitte manuell prüfen.";
        }
        showStatus(msg, shapeWarning ? "warning" : "success");

        // Formular zurücksetzen
        setTimeout(function () {
            document.getElementById("fields-container").innerHTML =
                "<p class=\"hint\">Fertig. Zum erneuten Einlesen: &#8635;</p>";
            document.getElementById("fill-btn").classList.add("hidden");
        }, 2500);

    } catch (err) {
        showStatus("Fehler: " + err.message, "error");
    } finally {
        fillBtn.disabled    = false;
        fillBtn.textContent = "✓  Ausfüllen";
    }
}

// ── Hilfs­funktion ────────────────────────────────────────────────────────────

function showStatus(msg, type) {
    var el = document.getElementById("status");
    if (!msg) {
        el.className = "status hidden";
        return;
    }
    el.className    = "status " + (type || "");
    el.textContent  = msg;
}
