(function (window, undefined) {
    "use strict";

    /*
     * Coastal Respiratory & Sleep Specialists
     * Xestro / ONLYOFFICE Clinical Letter Formatter v19
     *
     * PURPOSE
     * -------
     * Formats the clinical material inserted by LyreBird into the Xestro
     * GP-letter template. Xestro-owned content is deliberately left alone:
     * letterhead, date, Re:/demographics, salutation, signature/author block,
     * provider number, cc, disclaimer and footer.
     *
     * v19 final layout (matched to the approved 11.5 pt preview):
     * - clinical body text is 11.5 pt; major clinical section headings are 12 pt bold
     * - true numbered diagnosis/problem lists align their numbers with section headings
     * - dash subpoints use the approved 30 pt hanging indent so wrapped lines align with subpoint text
     * - investigations are compact italic bullet points with bold investigation labels
     * - blank paragraphs between consecutive investigations are removed
     * - Follow-up is bold, unnumbered, and aligned with numbered problem text
     * - optional/missing sections (including Investigations and Smoking status) require no special handling
     * - conservative pagination; Xestro-owned content remains untouched
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

                // Formatting dimensions in twips (20 twips = 1 point).
                // These values reproduce the approved final 11.5 pt preview.
                var BODY_FONT_SIZE = 23;             // 11.5 pt (ONLYOFFICE uses half-points)
                var HEADING_FONT_SIZE = 24;          // 12 pt
                var MAIN_LIST_TEXT_INDENT = 360;     // 18 pt: text position; number starts at left margin
                var SUBPOINT_TEXT_INDENT = 600;      // 30 pt: approved final subpoint text position
                var SUBPOINT_HANG = 180;             // 9 pt: dash sits at 21 pt; wrapped text starts at 30 pt
                var INVESTIGATION_TEXT_INDENT = 720; // 36 pt
                var INVESTIGATION_HANG = 360;        // bullet at 18 pt, text at 36 pt

                // LyreBird headings are accepted with or without trailing colons.
                var SECTION_HEADINGS = {
                    "HISTORY OF PRESENTING COMPLAINT": true,
                    "HISTORY OF PRESENTING ILLNESS": true,
                    "HOPC": true,
                    "INTERVAL HISTORY": true,
                    "REASON FOR REVIEW": true,
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
                    "FAMILY / SOCIAL HISTORY": true,
                    "FAMILY/SOCIAL HISTORY": true,
                    "TRAVEL HISTORY": true,
                    "OCCUPATION": true,
                    "MEDICATIONS": true,
                    "CURRENT MEDICATIONS": true,
                    "ADVERSE DRUG REACTIONS": true,
                    "ALLERGIES": true,
                    "ALLERGIES AND ADR": true,
                    "ALLERGIES & ADR": true,
                    "INVESTIGATIONS": true,
                    "EXAMINATION": true,
                    "EXAMINATION FINDINGS": true,
                    "PHYSICAL EXAMINATION": true,
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

                // Only diagnosis-style medical-history sections are numbered.
                // HOPC / interval history / respiratory / sleep / social sections remain prose.
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

                function isSalutation(text) {
                    return /^Dear\s+.+[,;:]?$/i.test((text || "").trim());
                }

                function looksLikeGenericClinicalHeading(text) {
                    var t = (text || "").trim();
                    if (!t || t.length > 70 || !/:$/.test(t)) return false;
                    var core = t.slice(0, -1).trim();
                    if (!/[A-Za-z]/.test(core)) return false;
                    return core === core.toUpperCase();
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
                    var numbering = doc.CreateNumbering("numbered");
                    var level = numbering.GetLevel(0);

                    // Keep a genuine numbered list, but put the number itself on the same
                    // left edge as the section heading. The item text begins 18 pt in.
                    level.SetCustomType("decimal", "%1.", "left");
                    level.SetSuff("tab");
                    var levelPr = level.GetParaPr();
                    levelPr.SetIndLeft(MAIN_LIST_TEXT_INDENT);
                    levelPr.SetIndFirstLine(-MAIN_LIST_TEXT_INDENT);
                    return numbering;
                }

                function makeInvestigationBullets() {
                    var numbering = doc.CreateNumbering("bullet");
                    var level = numbering.GetLevel(0);
                    level.SetSuff("tab");
                    var levelPr = level.GetParaPr();
                    levelPr.SetIndLeft(INVESTIGATION_TEXT_INDENT);
                    levelPr.SetIndFirstLine(-INVESTIGATION_HANG);
                    return numbering;
                }

                function applyMainNumbering(p, numbering) {
                    p.SetNumbering(numbering.GetLevel(0));
                    // Direct paragraph indents make the alignment deterministic even if
                    // ONLYOFFICE retains an older list indent on a re-formatted paragraph.
                    p.SetIndLeft(MAIN_LIST_TEXT_INDENT);
                    p.SetIndFirstLine(-MAIN_LIST_TEXT_INDENT);
                    p.SetSpacingBefore(30, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(false);
                    p.SetBold(true);
                    p.SetItalic(false);
                    p.SetFontSize(BODY_FONT_SIZE); // 11.5 pt
                    p.SetHighlight("none");
                    p.SetKeepLines(true);
                }

                function formatSubPoint(p) {
                    // Hanging indent: the dash remains at 36 pt, while wrapped lines start
                    // under the first word (e.g. "years" aligns with "Brother").
                    p.SetIndLeft(SUBPOINT_TEXT_INDENT);
                    p.SetIndFirstLine(-SUBPOINT_HANG);
                    p.SetSpacingBefore(0, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(true);
                    p.SetBold(false);
                    p.SetItalic(false);
                    p.SetFontSize(BODY_FONT_SIZE); // 11.5 pt
                    p.SetHighlight("none");
                    p.SetKeepLines(true);
                    p.SetWidowControl(true);
                }

                function formatUnnumberedHistoryLine(p) {
                    p.SetIndLeft(0);
                    p.SetSpacingBefore(80, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(false);
                    p.SetBold(false);
                    p.SetItalic(false);
                    p.SetFontSize(BODY_FONT_SIZE); // 11.5 pt
                    p.SetHighlight("none");
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
                    p.SetFontSize(HEADING_FONT_SIZE); // 12 pt section headings
                    p.SetHighlight("none");
                    p.SetIndLeft(0);
                    p.SetIndFirstLine(0);
                    p.SetSpacingBefore(80, false);
                    p.SetSpacingAfter(20, false);
                    p.SetContextualSpacing(false);
                    p.SetKeepLines(true);
                    p.SetKeepNext(true);
                }

                function formatFollowUpHeading(p) {
                    // Follow-up is a plan-level subheading, not a numbered clinical problem.
                    // Align it with the text of numbered problems (18 pt from the left).
                    p.SetBold(true);
                    p.SetItalic(false);
                    p.SetFontSize(BODY_FONT_SIZE); // 11.5 pt
                    p.SetHighlight("none");
                    p.SetIndLeft(MAIN_LIST_TEXT_INDENT);
                    p.SetIndFirstLine(0);
                    p.SetSpacingBefore(60, false); // 3 pt
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(false);
                    p.SetKeepLines(true);
                    p.SetKeepNext(true);
                }

                function formatFollowUpBody(p) {
                    p.SetIndLeft(SUBPOINT_TEXT_INDENT);
                    p.SetIndFirstLine(0);
                    p.SetFontSize(BODY_FONT_SIZE);
                    p.SetHighlight("none");
                    p.SetSpacingBefore(0, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(true);
                    p.SetKeepLines(true);
                    p.SetWidowControl(true);
                }

                function formatBody(p) {
                    // General LyreBird clinical prose. Preserve inline emphasis but normalise
                    // the approved body size and paragraph geometry.
                    p.SetFontSize(BODY_FONT_SIZE);
                    p.SetHighlight("none");
                    p.SetIndLeft(0);
                    p.SetIndFirstLine(0);
                    p.SetSpacingBefore(0, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(false);
                    p.SetWidowControl(true);
                }

                function formatInvestigation(p, bulletNumbering) {
                    // Compact bullet list. The full investigation line is italic; the
                    // investigation/provider/date label through the first colon is bold italic.
                    p.SetNumbering(bulletNumbering.GetLevel(0));
                    p.SetIndLeft(INVESTIGATION_TEXT_INDENT);
                    p.SetIndFirstLine(-INVESTIGATION_HANG);
                    p.SetItalic(true);
                    p.SetBold(false);
                    p.SetFontSize(BODY_FONT_SIZE); // 11.5 pt
                    p.SetHighlight("none");
                    p.SetSpacingBefore(0, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(true);
                    p.SetKeepLines(true);
                    p.SetWidowControl(true);

                    var t = cleanText(p);
                    var c = t.indexOf(":");
                    if (c >= 0) {
                        var r = p.GetRange(0, c + 1);
                        if (r) {
                            r.SetBold(true);
                            r.SetItalic(true);
                        }
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
                var investigationBulletNumbering = null;
                var investigationCount = 0;
                var investigationBlankParas = [];
                var clinicalFormattingStarted = false;
                var afterSalutation = false;

                for (var i = 0; i < paragraphs.length; i++) {
                    var p = paragraphs[i];
                    var text = cleanText(p);

                    if (!text) {
                        // Remove only blank paragraphs that sit BETWEEN consecutive
                        // investigation items. Keep the blank space after the final
                        // investigation before the narrative section.
                        if (mode === "investigations" && investigationCount > 0) {
                            var nextInvestigationText = nextNonEmptyText(i);
                            if (looksLikeInvestigation(nextInvestigationText)) {
                                investigationBlankParas.push(p);
                            }
                        }
                        continue;
                    }

                    if (isSalutation(text)) {
                        // The salutation itself belongs to Xestro. The following paragraph is
                        // the start of LyreBird's clinical body, even if no optional section
                        // headings are present.
                        afterSalutation = true;
                        mode = "";
                        clinicalFormattingStarted = false;
                        continue;
                    }

                    // Hard stop: never touch Xestro's sign-off / author / footer region.
                    if (isXestroBoundary(text)) {
                        mode = "";
                        clinicalFormattingStarted = false;
                        afterSalutation = false;
                        continue;
                    }

                    if (afterSalutation) {
                        clinicalFormattingStarted = true;
                        afterSalutation = false;
                    }

                    var h = normHeading(text);

                    // Follow-up is a PLAN subheading, not a separate major section.
                    if (isFollowUpHeading(text) && clinicalFormattingStarted) {
                        formatFollowUpHeading(p);
                        mode = "followup";
                        continue;
                    }

                    if (isHeading(text) || (clinicalFormattingStarted && looksLikeGenericClinicalHeading(text))) {
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
                            investigationBulletNumbering = makeInvestigationBullets();
                        } else if (ASSESSMENT_HEADINGS[h]) {
                            mode = "assessment";
                            assessmentNumbering = makeNumbering();
                        } else if (PLAN_HEADINGS[h]) {
                            mode = "plan";
                            planNumbering = makeNumbering();
                        }
                        continue;
                    }

                    // Before the clinical body begins, leave Xestro's wrapper untouched.
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

                    if (mode === "followup") {
                        if (isSubPoint(text)) {
                            formatSubPoint(p);
                        } else {
                            formatFollowUpBody(p);
                        }
                        continue;
                    }

                    if (mode === "investigations") {
                        if (looksLikeInvestigation(text)) {
                            formatInvestigation(p, investigationBulletNumbering);
                            investigationCount++;
                            continue;
                        }

                        // The first ordinary paragraph after INVESTIGATIONS marks the end
                        // of the investigation list. This also handles an intentionally empty
                        // Investigations section without breaking subsequent formatting.
                        mode = "";
                        formatBody(p);
                        continue;
                    }

                    // General clinical prose in optional or unrecognised sections.
                    if (clinicalFormattingStarted) {
                        formatBody(p);
                    }
                }

                // Delete blank paragraphs between investigation bullets in reverse order.
                // ApiParagraph.Delete is supported in current ONLYOFFICE builds. A tiny-line
                // fallback prevents a visible blank line on older builds.
                for (var b = investigationBlankParas.length - 1; b >= 0; b--) {
                    var bp = investigationBlankParas[b];
                    try {
                        if (bp && typeof bp.Delete === "function") {
                            bp.Delete();
                        } else if (bp) {
                            bp.SetFontSize(2);
                            bp.SetSpacingBefore(0, false);
                            bp.SetSpacingAfter(0, false);
                            if (typeof bp.SetSpacingLine === "function") bp.SetSpacingLine(20, "exact");
                        }
                    } catch (ignoreDeleteError) {
                        try {
                            bp.SetFontSize(2);
                            bp.SetSpacingBefore(0, false);
                            bp.SetSpacingAfter(0, false);
                            if (typeof bp.SetSpacingLine === "function") bp.SetSpacingLine(20, "exact");
                        } catch (ignoreFallbackError) {}
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

                // Match the approved preview: no forced paragraph-after spacing at
                // section ends; the source/template blank paragraphs provide separation.
                addSpacingAfterSection(HISTORY_HEADINGS, 0);
                addSpacingAfterSection(ASSESSMENT_HEADINGS, 0);
                addSpacingAfterSection(PLAN_HEADINGS, 0);

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
        setStatus("Coastal LyreBird Formatter v19 loaded — click Format current letter.", false);
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };
})(window, undefined);
