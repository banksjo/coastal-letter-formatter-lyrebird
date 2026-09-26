(function (window, undefined) {
    "use strict";

    /*
     * Coastal Respiratory & Sleep Specialists
     * Xestro / ONLYOFFICE Clinical Letter Formatter v26 XESTRO COMPATIBILITY FIX
     *
     * PURPOSE
     * -------
     * Formats the clinical material inserted by LyreBird into the Xestro
     * GP-letter template. Xestro-owned content is deliberately left alone:
     * letterhead, date, Re:/demographics, salutation, signature/author block,
     * provider number, cc, disclaimer and footer.
     *
     * v26 layout (measured directly from the user-corrected Greg DOCX):
     * - Arial 11 pt black justified clinical body; Arial 12 pt black bold-underlined major headings
     * - one blank line BEFORE major section headings; NO blank line AFTER headings
     * - headings have 0 pt paragraph-after spacing
     * - the opening LyreBird summary paragraph has 12 pt paragraph-before spacing
     * - numbered diagnoses/problems align number at the section left edge; item text begins 18 pt in
     * - dash subpoints use 36 pt text indent with a 9 pt hanging indent
     * - investigations are compact italic bullets with bold-italic investigation labels
     * - later numbered items have 1.5 pt before; subpoints have no extra paragraph spacing
     * - Smoking status/history remains unnumbered; Follow-up remains bold and unnumbered
     * - optional/missing sections remain safe; consecutive blank clinical lines collapse to one
     * - Xestro-owned salutation/sign-off/signature/footer remain untouched
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
                // These values reproduce the approved v21 typography plus the user-approved
                // Michelle section/line-spacing pattern.
                var BODY_FONT_SIZE = 22;             // 11 pt (ONLYOFFICE uses half-points)
                var HEADING_FONT_SIZE = 24;          // 12 pt
                var MAIN_LIST_TEXT_INDENT = 360;     // 18 pt: text position; number starts at left margin
                var SUBPOINT_TEXT_INDENT = 720;      // 36 pt: matches the slightly deeper indent in edited output 46
                var SUBPOINT_HANG = 180;             // 9 pt: dash sits at 27 pt; wrapped text starts at 36 pt
                var INVESTIGATION_TEXT_INDENT = 720; // 36 pt
                var INVESTIGATION_HANG = 360;        // bullet at 18 pt, text at 36 pt
                var BLANK_LINE_SPACING = 240;          // single-spaced blank paragraph (1.0 line)
                var HEADING_GAP_BEFORE = 80;           // 4 pt when a blank line already precedes later headings
                var HEADING_NO_BLANK_BEFORE = 320;     // 16 pt if LyreBird omits the normal pre-heading blank
                var FIRST_HEADING_NO_BLANK_BEFORE = 240; // 12 pt if the first heading has no source blank
                var FIRST_CLINICAL_PROSE_GAP = 240;    // 12 pt before the opening LyreBird summary paragraph
                var POST_INVESTIGATION_PROSE_GAP = 240; // 12 pt before narrative after investigations

                // Xestro may run an older ONLYOFFICE build than the current public API.
                // Optional formatting calls are therefore feature-detected so one unsupported
                // method cannot abort the whole formatter.
                function safeCall(obj, method, args) {
                    try {
                        if (obj && typeof obj[method] === "function") {
                            return obj[method].apply(obj, args || []);
                        }
                    } catch (ignoreCompatibilityError) {}
                    return null;
                }

                function safeBlack(p) {
                    try {
                        if (p && typeof p.SetColor === "function" && typeof Api.HexColor === "function") {
                            p.SetColor(Api.HexColor("#000000"));
                        }
                    } catch (ignoreColorCompatibilityError) {}
                }

                // Canonical display text for recognised section headings.
                // This allows LyreBird to output ALL CAPS, Title Case or a trailing colon
                // while the final letter is presented consistently.
                var CANONICAL_HEADINGS = {
                    "HISTORY OF PRESENTING COMPLAINT": "History of Presenting Complaint",
                    "HISTORY OF PRESENTING ILLNESS": "History of Presenting Illness",
                    "HOPC": "HOPC",
                    "INTERVAL HISTORY": "Interval History",
                    "REASON FOR REVIEW": "Reason for Review",
                    "BACKGROUND MEDICAL HISTORY": "Background Medical History",
                    "PAST MEDICAL HISTORY": "Past Medical History",
                    "MEDICAL HISTORY": "Medical History",
                    "RESPIRATORY HISTORY": "Respiratory History",
                    "SLEEP HISTORY": "Sleep History",
                    "SMOKING HISTORY": "Smoking History",
                    "SOCIAL HISTORY": "Social History",
                    "SOCIAL & EXPOSURE HISTORY": "Social & Exposure History",
                    "SOCIAL AND EXPOSURE HISTORY": "Social and Exposure History",
                    "OCCUPATIONAL HISTORY": "Occupational History",
                    "OCCUPATIONAL & EXPOSURE HISTORY": "Occupational & Exposure History",
                    "OCCUPATIONAL AND EXPOSURE HISTORY": "Occupational and Exposure History",
                    "EXPOSURE HISTORY": "Exposure History",
                    "FAMILY HISTORY": "Family History",
                    "FAMILY / SOCIAL HISTORY": "Family / Social History",
                    "FAMILY/SOCIAL HISTORY": "Family/Social History",
                    "TRAVEL HISTORY": "Travel History",
                    "OCCUPATION": "Occupation",
                    "MEDICATIONS": "Medications",
                    "CURRENT MEDICATIONS": "Current Medications",
                    "ADVERSE DRUG REACTIONS": "Adverse Drug Reactions",
                    "ALLERGIES": "Allergies",
                    "ALLERGIES AND ADR": "Allergies and ADR",
                    "ALLERGIES & ADR": "Allergies & ADR",
                    "INVESTIGATIONS": "Investigations",
                    "EXAMINATION": "Examination",
                    "EXAMINATION FINDINGS": "Examination Findings",
                    "PHYSICAL EXAMINATION": "Physical Examination",
                    "CURRENT ASSESSMENT": "Current Assessment",
                    "ASSESSMENT": "Assessment",
                    "ASSESSMENT / ISSUES": "Assessment / Issues",
                    "ASSESSMENT/ISSUES": "Assessment/Issues",
                    "ASSESSMENT / KEY ISSUES": "Assessment / Key Issues",
                    "ASSESSMENT/KEY ISSUES": "Assessment/Key Issues",
                    "ASSESSMENT AND ISSUES": "Assessment and Issues",
                    "ASSESSMENT & ISSUES": "Assessment & Issues",
                    "IMPRESSION": "Impression",
                    "IMPRESSION / DIAGNOSIS": "Impression / Diagnosis",
                    "IMPRESSION/DIAGNOSIS": "Impression/Diagnosis",
                    "DIAGNOSIS": "Diagnosis",
                    "CURRENT PROBLEMS": "Current Problems",
                    "PLAN": "Plan",
                    "MANAGEMENT PLAN": "Management Plan"
                };

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

                function canonicalHeadingText(text) {
                    var h = normHeading(text);
                    return CANONICAL_HEADINGS[h] || null;
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
                    return /^(Smoking status|Smoking history|Smoking|Tobacco use|Pack[- ]?years?|Vaping history|E-cigarette use)\s*:/i.test(t);
                }

                function isXestroBoundary(text) {
                    var t = (text || "").trim();
                    return (
                        /^Yours sincerely[,;]?$/i.test(t) ||
                        /^Kind regards[,;]?$/i.test(t) ||
                        /^Regards[,;]?$/i.test(t) ||
                        /^Dr\.?\s+Jonathan\s+Banks\b/i.test(t) ||
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

                function normaliseClinicalText(p) {
                    if (!p) return;
                    safeCall(p, "SetFontFamily", ["Arial"]);
                    safeBlack(p);
                    safeCall(p, "SetHighlight", ["none"]);
                    safeCall(p, "SetSpacingLine", [240, "auto"]);
                    safeCall(p, "SetJc", ["both"]);
                    // These methods are known to exist in the older Xestro build, but still
                    // guard them so a version mismatch can never stop the formatter.
                    safeCall(p, "SetWidowControl", [true]);
                }

                function makeNumbering() {
                    // Use ONLYOFFICE's built-in numbered list. Custom list-level APIs were
                    // introduced later and can fail in the older editor embedded by Xestro.
                    return doc.CreateNumbering("numbered");
                }

                function makeInvestigationBullets() {
                    return doc.CreateNumbering("bullet");
                }

                function applyMainNumbering(p, numbering, isFirstItem) {
                    normaliseClinicalText(p);
                    p.SetNumbering(numbering.GetLevel(0));
                    // Direct paragraph indents make the alignment deterministic even if
                    // ONLYOFFICE retains an older list indent on a re-formatted paragraph.
                    p.SetIndLeft(MAIN_LIST_TEXT_INDENT);
                    safeCall(p, "SetIndFirstLine", [-MAIN_LIST_TEXT_INDENT]);
                    p.SetSpacingBefore(isFirstItem ? 0 : 30, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(false);
                    p.SetBold(true);
                    p.SetItalic(false);
                    safeCall(p, "SetUnderline", ["none"]);
                    safeCall(p, "SetFontSize", [BODY_FONT_SIZE]); // 11 pt
                    safeCall(p, "SetHighlight", ["none"]);
                    p.SetKeepLines(true);
                }

                function formatSubPoint(p) {
                    // Proper hanging indent: no literal leading spaces are inserted.
                    // The dash sits at 27 pt and all wrapped lines begin at 36 pt,
                    // matching the edited output while keeping multi-line subpoints aligned.
                    normaliseClinicalText(p);
                    p.SetIndLeft(SUBPOINT_TEXT_INDENT);
                    safeCall(p, "SetIndFirstLine", [-SUBPOINT_HANG]);
                    p.SetSpacingBefore(0, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(true);
                    p.SetBold(false);
                    p.SetItalic(false);
                    safeCall(p, "SetUnderline", ["none"]);
                    safeCall(p, "SetFontSize", [BODY_FONT_SIZE]); // 11 pt
                    safeCall(p, "SetHighlight", ["none"]);
                    p.SetKeepLines(true);
                    p.SetWidowControl(true);
                }

                function formatUnnumberedHistoryLine(p) {
                    normaliseClinicalText(p);
                    p.SetIndLeft(0);
                    p.SetSpacingBefore(80, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(false);
                    p.SetBold(false);
                    p.SetItalic(false);
                    safeCall(p, "SetUnderline", ["none"]);
                    safeCall(p, "SetFontSize", [BODY_FONT_SIZE]); // 11 pt
                    safeCall(p, "SetHighlight", ["none"]);
                    p.SetKeepLines(true);
                    var t = cleanText(p);
                    var c = t.indexOf(":");
                    if (c >= 0 && typeof p.GetRange === "function") {
                        var r = safeCall(p, "GetRange", [0, c + 1]);
                        if (r) safeCall(r, "SetBold", [true]);
                    }
                }

                function formatHeading(p, isFirstMajorHeading, hasBlankBefore, hasBlankAfter) {
                    normaliseClinicalText(p);
                    p.SetBold(true);
                    p.SetItalic(false);
                    safeCall(p, "SetUnderline", ["single"]);
                    safeCall(p, "SetFontSize", [HEADING_FONT_SIZE]); // 12 pt section headings
                    p.SetIndLeft(0);
                    safeCall(p, "SetIndFirstLine", [0]);

                    // Exact rhythm from the corrected Greg file:
                    // preserve one blank line BEFORE a major heading, remove any blank line
                    // AFTER it, and place the first content line immediately under the heading.
                    if (hasBlankBefore) {
                        p.SetSpacingBefore(isFirstMajorHeading ? 0 : HEADING_GAP_BEFORE, false);
                    } else {
                        p.SetSpacingBefore(isFirstMajorHeading ? FIRST_HEADING_NO_BLANK_BEFORE : HEADING_NO_BLANK_BEFORE, false);
                    }
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(false);
                    p.SetKeepLines(true);
                    p.SetKeepNext(true);
                }

                function formatFollowUpHeading(p) {
                    normaliseClinicalText(p);
                    // Follow-up is a plan-level subheading, not a numbered clinical problem.
                    // Align it with the text of numbered problems (18 pt from the left).
                    p.SetBold(true);
                    p.SetItalic(false);
                    safeCall(p, "SetUnderline", ["none"]);
                    safeCall(p, "SetFontSize", [BODY_FONT_SIZE]); // 11 pt
                    safeCall(p, "SetHighlight", ["none"]);
                    p.SetIndLeft(MAIN_LIST_TEXT_INDENT);
                    safeCall(p, "SetIndFirstLine", [0]);
                    p.SetSpacingBefore(60, false); // 3 pt
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(false);
                    p.SetKeepLines(true);
                    p.SetKeepNext(true);
                }

                function formatFollowUpBody(p) {
                    normaliseClinicalText(p);
                    p.SetIndLeft(SUBPOINT_TEXT_INDENT);
                    safeCall(p, "SetIndFirstLine", [0]);
                    safeCall(p, "SetFontSize", [BODY_FONT_SIZE]);
                    safeCall(p, "SetHighlight", ["none"]);
                    p.SetSpacingBefore(0, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(true);
                    p.SetKeepLines(true);
                    p.SetWidowControl(true);
                }

                function formatBody(p, spacingBeforeTwips) {
                    normaliseClinicalText(p);
                    // General LyreBird clinical prose. Preserve inline emphasis but normalise
                    // the approved body size and paragraph geometry.
                    safeCall(p, "SetFontSize", [BODY_FONT_SIZE]);
                    safeCall(p, "SetHighlight", ["none"]);
                    p.SetIndLeft(0);
                    safeCall(p, "SetIndFirstLine", [0]);
                    p.SetSpacingBefore(spacingBeforeTwips || 0, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(false);
                    p.SetWidowControl(true);
                }

                function formatInvestigation(p, bulletNumbering) {
                    normaliseClinicalText(p);
                    // Compact bullet list. The full investigation line is italic; the
                    // investigation/provider/date label through the first colon is bold italic.
                    p.SetNumbering(bulletNumbering.GetLevel(0));
                    p.SetIndLeft(INVESTIGATION_TEXT_INDENT);
                    safeCall(p, "SetIndFirstLine", [-INVESTIGATION_HANG]);
                    p.SetItalic(true);
                    p.SetBold(false);
                    safeCall(p, "SetUnderline", ["none"]);
                    safeCall(p, "SetFontSize", [BODY_FONT_SIZE]); // 11 pt
                    safeCall(p, "SetHighlight", ["none"]);
                    p.SetSpacingBefore(0, false);
                    p.SetSpacingAfter(0, false);
                    p.SetContextualSpacing(true);
                    p.SetKeepLines(true);
                    p.SetWidowControl(true);

                    var t = cleanText(p);
                    var c = t.indexOf(":");
                    if (c >= 0 && typeof p.GetRange === "function") {
                        var r = safeCall(p, "GetRange", [0, c + 1]);
                        if (r) {
                            safeCall(r, "SetBold", [true]);
                            safeCall(r, "SetItalic", [true]);
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

                function hasImmediateBlankBefore(index) {
                    return index > 0 && !cleanText(paragraphs[index - 1]);
                }

                function hasImmediateBlankAfter(index) {
                    return index + 1 < paragraphs.length && !cleanText(paragraphs[index + 1]);
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
                var blankParasToDelete = [];
                var previousClinicalWasBlank = false;

                function queueBlankDelete(bp) {
                    if (!bp) return;
                    if (blankParasToDelete.indexOf(bp) < 0) blankParasToDelete.push(bp);
                }
                var historyItemCount = 0;
                var assessmentItemCount = 0;
                var planItemCount = 0;
                var clinicalFormattingStarted = false;
                var afterSalutation = false;
                var majorHeadingCount = 0;
                var firstClinicalContentSeen = false;

                for (var i = 0; i < paragraphs.length; i++) {
                    var p = paragraphs[i];
                    var text = cleanText(p);

                    if (!text) {
                        // Within the LyreBird clinical area, preserve one intentional blank line,
                        // but collapse runs of two or more blank paragraphs. This makes the
                        // formatter tolerant of small LyreBird spacing variations while keeping
                        // the user-approved Michelle vertical rhythm.
                        if (clinicalFormattingStarted || afterSalutation) {
                            try {
                                p.SetSpacingBefore(0, false);
                                p.SetSpacingAfter(0, false);
                                safeCall(p, "SetSpacingLine", [BLANK_LINE_SPACING, "auto"]);
                            } catch (ignoreBlankFormattingError) {}

                            if (previousClinicalWasBlank) {
                                queueBlankDelete(p);
                            } else {
                                previousClinicalWasBlank = true;
                            }
                        }

                        // Blank paragraphs BETWEEN consecutive investigation items are removed
                        // entirely so the investigation list stays compact. Keep one blank after
                        // the final investigation before narrative prose.
                        if (mode === "investigations" && investigationCount > 0) {
                            var nextInvestigationText = nextNonEmptyText(i);
                            if (looksLikeInvestigation(nextInvestigationText)) {
                                queueBlankDelete(p);
                            }
                        }
                        continue;
                    }

                    if (isSalutation(text)) {
                        // The salutation itself belongs to Xestro. The following paragraph is
                        // the start of LyreBird's clinical body, even if no optional section
                        // headings are present.
                        afterSalutation = true;
                        previousClinicalWasBlank = false;
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

                    var isFirstClinicalContent = false;
                    if (afterSalutation) {
                        clinicalFormattingStarted = true;
                        afterSalutation = false;
                        if (!firstClinicalContentSeen) {
                            isFirstClinicalContent = true;
                            firstClinicalContentSeen = true;
                        }
                    }
                    if (clinicalFormattingStarted) previousClinicalWasBlank = false;

                    var h = normHeading(text);

                    // Follow-up is a PLAN subheading, not a separate major section.
                    if (isFollowUpHeading(text) && clinicalFormattingStarted) {
                        formatFollowUpHeading(p);
                        mode = "followup";
                        continue;
                    }

                    if (isHeading(text) || (clinicalFormattingStarted && looksLikeGenericClinicalHeading(text))) {
                        clinicalFormattingStarted = true;

                        // For recognised headings, replace the visible text with the canonical
                        // Title Case form used in output 46. ApiParagraph.SetText is used only
                        // on the heading paragraph itself; list/body paragraphs are untouched.
                        var canonicalHeading = canonicalHeadingText(text);
                        if (canonicalHeading && typeof p.SetText === "function") {
                            safeCall(p, "SetText", [canonicalHeading]);
                        }

                        var blankAfterHeading = hasImmediateBlankAfter(i);
                        formatHeading(p, majorHeadingCount === 0, hasImmediateBlankBefore(i), blankAfterHeading);
                        // User-corrected layout: no blank paragraph after major headings.
                        if (blankAfterHeading && i + 1 < paragraphs.length) {
                            queueBlankDelete(paragraphs[i + 1]);
                        }
                        majorHeadingCount++;

                        mode = "";
                        investigationCount = 0;

                        if (HISTORY_HEADINGS[h]) {
                            mode = "history";
                            historyNumbering = makeNumbering();
                            historyItemCount = 0;
                        } else if (h === "INVESTIGATIONS") {
                            mode = "investigations";
                            investigationBulletNumbering = makeInvestigationBullets();
                        } else if (ASSESSMENT_HEADINGS[h]) {
                            mode = "assessment";
                            assessmentNumbering = makeNumbering();
                            assessmentItemCount = 0;
                        } else if (PLAN_HEADINGS[h]) {
                            mode = "plan";
                            planNumbering = makeNumbering();
                            planItemCount = 0;
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
                            applyMainNumbering(p, historyNumbering, historyItemCount === 0);
                            historyItemCount++;
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
                            applyMainNumbering(p, assessmentNumbering, assessmentItemCount === 0);
                            assessmentItemCount++;
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
                            applyMainNumbering(p, planNumbering, planItemCount === 0);
                            planItemCount++;
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
                        formatBody(p, POST_INVESTIGATION_PROSE_GAP);
                        continue;
                    }

                    // General clinical prose in optional or unrecognised sections.
                    if (clinicalFormattingStarted) {
                        formatBody(p, isFirstClinicalContent ? FIRST_CLINICAL_PROSE_GAP : 0);
                    }
                }

                // Delete redundant clinical blanks (including blanks between investigation
                // bullets) in reverse order. ApiParagraph.Delete is supported in current
                // ONLYOFFICE builds. A tiny-line fallback prevents a visible gap on older builds.
                for (var b = blankParasToDelete.length - 1; b >= 0; b--) {
                    var bp = blankParasToDelete[b];
                    try {
                        if (bp && typeof bp.Delete === "function") {
                            bp.Delete();
                        } else if (bp) {
                            safeCall(bp, "SetFontSize", [2]);
                            bp.SetSpacingBefore(0, false);
                            bp.SetSpacingAfter(0, false);
                            if (typeof bp.SetSpacingLine === "function") safeCall(bp, "SetSpacingLine", [20, "exact"]);
                        }
                    } catch (ignoreDeleteError) {
                        try {
                            safeCall(bp, "SetFontSize", [2]);
                            bp.SetSpacingBefore(0, false);
                            bp.SetSpacingAfter(0, false);
                            if (typeof bp.SetSpacingLine === "function") safeCall(bp, "SetSpacingLine", [20, "exact"]);
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
        setStatus("Coastal LyreBird Formatter v26 XESTRO COMPATIBILITY FIX loaded — click Format current letter.", false);
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };
})(window, undefined);
