(function (window, undefined) {
    "use strict";

    /*
     * Coastal Respiratory & Sleep Specialists
     * Xestro / ONLYOFFICE Clinical Letter Formatter v17
     *
     * PURPOSE
     * -------
     * Formats the clinical material inserted by LyreBird into the Xestro
     * GP-letter template. Xestro-owned content is deliberately left alone:
     * letterhead, date, Re:/demographics, salutation, signature/author block,
     * provider number, cc, disclaimer and footer.
     *
     * v17 professional layout:
     * - restores the reliable v15 numbering structure
     * - clear hierarchy: bold numbered diagnoses/problems, indented dash subpoints
     * - conservative pagination: headings stay with next paragraph only
     * - investigation label bold, result text normal (not whole-paragraph italics)
     * - Follow-up is an unnumbered bold subheading inside PLAN
     * - Smoking status remains unnumbered
     * - no text rewriting while the document is being formatted
     */

    function setStatus(text, isError) {
        var el = document.getElementById("status");
        if (!el) return;
        el.textContent = text;
        el.style.color = isError ? "#9b1c1c" : "#333333";
    }

    window.formatCoastalLetter = function () {
        var btn = document.getElementById("formatBtn");
        if (btn) btn.disabled = true;
        setStatus("Formatting clinical content…", false);

        window.Asc.plugin.callCommand(function () {
            try {
                var doc = Api.GetDocument();
                var paragraphs = doc.GetAllParagraphs();

                // Formatting dimensions in twips.
                var SUBPOINT_INDENT = 720;        // ~1.27 cm
                var INVESTIGATION_INDENT = 360;   // subtle ~0.64 cm

                // LyreBird headings are accepted with or without trailing colons.
                var SECTION_HEADINGS = {
                    "BACKGROUND MEDICAL HISTORY": true,
                    "PAST MEDICAL HISTORY": true,
                    "MEDICAL HISTORY": true,
                    "RESPIRATORY HISTORY": true,
                    "SLEEP HISTORY": true,
                    "SMOKING HISTORY": true,
                    "SOCIAL HISTORY": true,
                    "SOCIAL & EXPOSURE HISTORY": true,
                    "SOCIAL AND EXPOSURE HISTORY": true,
                    "OCCUPATIONAL HISTORY": true,
                    "OCCUPATIONAL & EXPOSURE HISTORY": true,
                    "OCCUPATIONAL AND EXPOSURE HISTORY": true,
                    "EXPOSURE HISTORY": true,
                    "FAMILY HISTORY": true,
                    "TRAVEL HISTORY": true,
                    "MEDICATIONS": true,
                    "CURRENT MEDICATIONS": true,
                    "ADVERSE DRUG REACTIONS": true,
                    "ALLERGIES": true,
                    "ALLERGIES AND ADR": true,
                    "ALLERGIES & ADR": true,
                    "INVESTIGATIONS": true,
                    "EXAMINATION": true,
                    "EXAMINATION FINDINGS": true,
                    "CURRENT ASSESSMENT": true,
                    "ASSESSMENT": true,
                    "ASSESSMENT / ISSUES": true,
                    "ASSESSMENT/ISSUES": true,
                    "ASSESSMENT / KEY ISSUES": true,
                    "ASSESSMENT/KEY ISSUES": true,
                    "ASSESSMENT AND ISSUES": true,
                    "ASSESSMENT & ISSUES": true,
                    "IMPRESSION": true,
                    "IMPRESSION / DIAGNOSIS": true,
                    "IMPRESSION/DIAGNOSIS": true,
                    "DIAGNOSIS": true,
                    "CURRENT PROBLEMS": true,
                    "PLAN": true,
                    "MANAGEMENT PLAN": true,
                    "FOLLOW-UP": true,
                    "FOLLOW UP": true
                };

                var HISTORY_HEADINGS = {
                    "BACKGROUND MEDICAL HISTORY": true,
                    "PAST MEDICAL HISTORY": true,
                    "MEDICAL HISTORY": true
                };

                var ASSESSMENT_HEADINGS = {
                    "CURRENT ASSESSMENT": true,
                    "ASSESSMENT": true,
                    "ASSESSMENT / ISSUES": true,
                    "ASSESSMENT/ISSUES": true,
                    "ASSESSMENT / KEY ISSUES": true,
                    "ASSESSMENT/KEY ISSUES": true,
                    "ASSESSMENT AND ISSUES": true,
                    "ASSESSMENT & ISSUES": true,
                    "IMPRESSION": true,
                    "IMPRESSION / DIAGNOSIS": true,
                    "IMPRESSION/DIAGNOSIS": true,
                    "DIAGNOSIS": true,
                    "CURRENT PROBLEMS": true
                };

                var PLAN_HEADINGS = {
                    "PLAN": true,
                    "MANAGEMENT PLAN": true
                };

                function cleanText(p) {
                    return (p.GetText({
                        "Numbering": false,
                        "Math": true,
                        "NewLineSeparator": "\n",
                        "TabSymbol": "\t"
                    }) || "").replace(/\u00A0/g, " ").trim();
                }

                function normHeading(text) {
                    return (text || "")
                        .replace(/\s+/g, " ")
                        .trim()
                        .replace(/:$/, "")
                        .toUpperCase();
                }

                function isHeading(text) {
                    return !!SECTION_HEADINGS[normHeading(text)];
                }

                function isFollowUpHeading(text) {
                    var h = normHeading(text);
                    return h === "FOLLOW-UP" || h === "FOLLOW UP";
                }

                function isSubPoint(text) {
                    return /^\s*[-–—•]\s+/.test(text || "");
                }

                function isUnnumberedHistoryLine(text) {
                    var t = (text || "").trim();
                    return /^(Smoking status|Smoking|Tobacco use|Pack[- ]?years?)\s*:/i.test(t);
                }

                function isXestroBoundary(text) {
                    var t = (text || "").trim();
                    return (
                        /^Yours sincerely[,;]?$/i.test(t) ||
                        /^Kind regards[,;]?$/i.test(t) ||
                        /^Regards[,;]?$/i.test(t) ||
                        /^Dr\s+Jonathan\s+Banks\b/i.test(t) ||
                        /^Resp\/Sleep\/Gen Med Phys\./i.test(t) ||
                        /^Coastal Respiratory & Sleep Specialists$/i.test(t) ||
                        /^Prov:/i.test(t) ||
                        /^cc:/i.test(t) ||
                        /^This letter may have been produced with the assistance of AI technology/i.test(t) ||
                        /^Coastal Respiratory and Sleep Services/i.test(t)
                    );
                }

                function looksLikeInvestigation(text) {
                    var t = (text || "").trim();

                    // Broad respiratory / sleep / general physician investigation vocabulary.
                    if (/^(RFTs?|PFTs?|Respiratory function tests?|Lung function|Spirometry|Plethysmography|DLCO|KCO|FeNO|Nasal nitric oxide|6[- ]?minute walk|6MWT|Walk test|Exercise oximetry|CT\b|HRCT\b|CTPA\b|CT chest\b|CT sinuses?\b|CT coronary\b|PET\b|PET\/CT\b|MRI\b|CXR\b|Chest X[- ]?ray|X[- ]?ray|Ultrasound\b|Echo\b|Echocardiogram\b|ECG\b|Holter\b|ABG\b|VBG\b|Blood\b|Blood tests?\b|FBC\b|UEC\b|LFTs?\b|CRP\b|ESR\b|IgE\b|IgG\b|Eosinophils?\b|Serology\b|ANA\b|ANCA\b|ENA\b|Biopsy\b|Lung biopsy\b|Histology\b|Pathology\b|Respiratory culture\b|Sputum\b|BAL\b|Bronchoscopy\b|EBUS\b|EBUS-TBNA\b|Sleep study\b|Home sleep study\b|Ambulatory home sleep study\b|PSG\b|Polysomnography\b|CPAP titration study\b|CPAP download\b|PAP download\b|Device download\b|Oximetry\b|Cystic fibrosis testing\b|Cystic fibrosis gene panel\b|Genetic testing\b)/i.test(t)) {
                        return true;
                    }

                    // Provider/date structured investigations: Investigation (provider, dd/mm/yy):
                    if (/^[A-Za-z][A-Za-z0-9 \-\/%&+().]{1,70}\s*\([^)]*,\s*\d{1,2}\/\d{1,2}\/\d{2,4}\)\s*:/i.test(t)) {
                        return true;
                    }

                    // Date-labelled tests such as "RFTs – 22/09/2026:".
                    if (/^[A-Za-z][A-Za-z0-9 \-\/%&+().]{1,70}\s*[–—-]\s*\d{1,2}\/\d{1,2}\/\d{2,4}\s*:/i.test(t)) {
                        return true;
                    }

                    return false;
                }

                function makeNumbering() {
                    return doc.CreateNumbering("numbered");
                }

                function applyMainNumbering(p, numbering) {
                    p.SetNumbering(numbering.GetLevel(0));
                    p.SetIndLeft(0);
                    p.SetSpacingBefore(60, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(false);
                    p.SetBold(true);
                    p.SetItalic(false);
                    p.SetKeepLines(true);
                }

                function formatSubPoint(p) {
                    p.SetIndLeft(SUBPOINT_INDENT);
                    p.SetSpacingBefore(0, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(true);
                    p.SetBold(false);
                    p.SetItalic(false);
                    p.SetKeepLines(true);
                    p.SetWidowControl(true);
                }

                function formatUnnumberedHistoryLine(p) {
                    p.SetIndLeft(0);
                    p.SetSpacingBefore(120, false);
                    p.SetSpacingAfter(120, false);
                    p.SetContextualSpacing(false);
                    p.SetBold(false);
                    p.SetItalic(false);
                    p.SetKeepLines(true);
                    var t = cleanText(p);
                    var c = t.indexOf(":");
                    if (c >= 0) {
                        var r = p.GetRange(0, c + 1);
                        if (r) r.SetBold(true);
                    }
                }

                function formatHeading(p) {
                    p.SetBold(true);
                    p.SetItalic(false);
                    p.SetIndLeft(0);
                    p.SetSpacingBefore(120, false);
                    p.SetSpacingAfter(40, false);
                    p.SetContextualSpacing(false);
                    p.SetKeepLines(true);
                    p.SetKeepNext(true);
                }

                function formatFollowUpHeading(p) {
                    p.SetBold(true);
                    p.SetItalic(false);
                    p.SetIndLeft(0);
                    p.SetSpacingBefore(120, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(false);
                    p.SetKeepLines(true);
                    p.SetKeepNext(true);
                }

                function formatInvestigation(p) {
                    p.SetItalic(false);
                    p.SetBold(false);
                    p.SetIndLeft(INVESTIGATION_INDENT);
                    p.SetSpacingBefore(0, false);
                    p.SetSpacingAfter(120, false);
                    p.SetContextualSpacing(false);
                    p.SetKeepLines(true);
                    p.SetWidowControl(true);

                    // Bold only the investigation label through the first colon.
                    var t = cleanText(p);
                    var c = t.indexOf(":");
                    if (c >= 0) {
                        var r = p.GetRange(0, c + 1);
                        if (r) r.SetBold(true);
                    }
                }

                function nextNonEmptyText(startIndex) {
                    for (var n = startIndex + 1; n < paragraphs.length; n++) {
                        var nt = cleanText(paragraphs[n]);
                        if (nt) return nt;
                    }
                    return "";
                }

                /*
                 * FIRST PASS: locate and format the clinical body.
                 */
                var mode = "";
                var historyNumbering = null;
                var assessmentNumbering = null;
                var planNumbering = null;
                var investigationCount = 0;
                var clinicalFormattingStarted = false;

                for (var i = 0; i < paragraphs.length; i++) {
                    var p = paragraphs[i];
                    var text = cleanText(p);

                    if (!text) continue;

                    // Hard stop: never touch Xestro's sign-off / author / footer region.
                    if (isXestroBoundary(text)) {
                        mode = "";
                        clinicalFormattingStarted = false;
                        continue;
                    }

                    var h = normHeading(text);

                    // Follow-up is a PLAN subheading, not a separate major section.
                    if (isFollowUpHeading(text) && clinicalFormattingStarted) {
                        formatFollowUpHeading(p);
                        mode = "plan";
                        if (!planNumbering) planNumbering = makeNumbering();
                        continue;
                    }

                    if (isHeading(text)) {
                        clinicalFormattingStarted = true;

                        // Do not rewrite paragraph text here. Changing paragraph text while
                        // iterating can invalidate ONLYOFFICE paragraph references and was the
                        // cause of the flattened v16 assessment/plan formatting.
                        formatHeading(p);

                        mode = "";
                        investigationCount = 0;

                        if (HISTORY_HEADINGS[h]) {
                            mode = "history";
                            historyNumbering = makeNumbering();
                        } else if (h === "INVESTIGATIONS") {
                            mode = "investigations";
                        } else if (ASSESSMENT_HEADINGS[h]) {
                            mode = "assessment";
                            assessmentNumbering = makeNumbering();
                        } else if (PLAN_HEADINGS[h]) {
                            mode = "plan";
                            planNumbering = makeNumbering();
                        }
                        continue;
                    }

                    // No recognised section = no formatting. This keeps the Xestro
                    // wrapper and ordinary narrative paragraphs untouched.
                    if (!clinicalFormattingStarted && !mode) {
                        continue;
                    }

                    if (mode === "history") {
                        if (isSubPoint(text)) {
                            formatSubPoint(p);
                        } else if (isUnnumberedHistoryLine(text)) {
                            formatUnnumberedHistoryLine(p);
                        } else {
                            applyMainNumbering(p, historyNumbering);
                            if (isSubPoint(nextNonEmptyText(i))) {
                                p.SetKeepNext(true);
                            }
                        }
                        continue;
                    }

                    if (mode === "assessment") {
                        if (isSubPoint(text)) {
                            formatSubPoint(p);
                        } else {
                            applyMainNumbering(p, assessmentNumbering);
                            if (isSubPoint(nextNonEmptyText(i))) {
                                p.SetKeepNext(true);
                            }
                        }
                        continue;
                    }

                    if (mode === "plan") {
                        if (isSubPoint(text)) {
                            formatSubPoint(p);
                        } else {
                            applyMainNumbering(p, planNumbering);
                            if (isSubPoint(nextNonEmptyText(i))) {
                                p.SetKeepNext(true);
                            }
                        }
                        continue;
                    }

                    if (mode === "investigations") {
                        if (looksLikeInvestigation(text)) {
                            formatInvestigation(p);
                            investigationCount++;
                            continue;
                        }

                        // After at least one investigation, the first ordinary paragraph
                        // is the HOPC/review narrative. Stop investigation styling.
                        if (investigationCount > 0) {
                            mode = "";
                        }
                    }
                }

                /*
                 * SECOND PASS: compact list spacing and visual separation.
                 */
                paragraphs = doc.GetAllParagraphs();

                function addSpacingAfterSection(headingMap, spacingTwips) {
                    var inSection = false;
                    var lastPara = null;

                    for (var x = 0; x < paragraphs.length; x++) {
                        var sp = paragraphs[x];
                        var st = cleanText(sp);
                        if (!st) continue;

                        var sh = normHeading(st);

                        if (headingMap[sh]) {
                            inSection = true;
                            lastPara = null;
                            continue;
                        }

                        // Follow-up remains within PLAN.
                        if (inSection && isFollowUpHeading(st)) {
                            continue;
                        }

                        if (inSection && isHeading(st)) {
                            if (lastPara) {
                                lastPara.SetSpacingAfter(spacingTwips, false);
                                lastPara.SetContextualSpacing(false);
                            }
                            inSection = false;
                            continue;
                        }

                        if (inSection && isXestroBoundary(st)) {
                            if (lastPara) {
                                lastPara.SetSpacingAfter(spacingTwips, false);
                                lastPara.SetContextualSpacing(false);
                            }
                            inSection = false;
                            continue;
                        }

                        if (inSection) {
                            // Preserve the paragraph-level spacing already applied above;
                            // only remember the final paragraph so the section can end cleanly.
                            lastPara = sp;
                        }
                    }

                    if (inSection && lastPara) {
                        lastPara.SetSpacingAfter(spacingTwips, false);
                        lastPara.SetContextualSpacing(false);
                    }
                }

                // One visual line after background medical history.
                addSpacingAfterSection(HISTORY_HEADINGS, 240);

                // Assessment and plan remain compact.
                addSpacingAfterSection(ASSESSMENT_HEADINGS, 120);
                addSpacingAfterSection(PLAN_HEADINGS, 120);

                /*
                 * THIRD PASS: conservative pagination control.
                 * Keep each major heading with the immediately following paragraph only.
                 * Do not chain multiple investigation paragraphs together: that caused
                 * the large awkward page movement seen in v16.
                 */
                paragraphs = doc.GetAllParagraphs();
                for (var q = 0; q < paragraphs.length; q++) {
                    var qp = paragraphs[q];
                    var qt = cleanText(qp);
                    if (!qt) continue;

                    if (isHeading(qt) || isFollowUpHeading(qt)) {
                        qp.SetKeepLines(true);
                        qp.SetKeepNext(true);
                    }

                    if (looksLikeInvestigation(qt)) {
                        qp.SetKeepLines(true);
                        qp.SetWidowControl(true);
                    }
                }

                return "OK";

            } catch (e) {
                return "ERROR: " + (e && e.message ? e.message : String(e));
            }
        }, true, true, function (result) {
            if (btn) btn.disabled = false;
            if (result && String(result).indexOf("ERROR:") === 0) {
                setStatus(result, true);
            } else {
                setStatus("Formatting complete. Review the letter, then save/send.", false);
            }
        });
    };

    window.Asc.plugin.init = function () {
        setStatus("Coastal LyreBird Formatter v17 loaded — click Format current letter.", false);
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };
})(window, undefined);
