# Patch: Insert Section 13 (Medical/Recovery) and renumber Conclusion to 14
$ErrorActionPreference = "Stop"
$docPath = "C:\Users\alforganym\OneDrive\1 - My Learning Path\1 - Long Term Projects\3 - SC Ploicy WorkFlow\2 - SC Policy Draft (Refined v1 - WIP).docx"

$wdStory = 6
$wdParagraph = 4
$wdMove = 0

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$word.ScreenUpdating = $false

$origUser = $word.UserName
$origInit = $word.UserInitials
$word.UserName = "Mahmoud AlforGany"
$word.UserInitials = "MA"

function Type-Heading2 { param($sel, $text)
    $sel.TypeParagraph(); $sel.Style = "Heading 2"; $sel.TypeText($text)
}
function Type-Body { param($sel, $text, $bold=$false)
    $sel.TypeParagraph(); $sel.Style = "Normal"
    if ($bold) { $sel.Font.Bold = $true; $sel.TypeText($text); $sel.Font.Bold = $false }
    else { $sel.TypeText($text) }
}
function Type-Bullet { param($sel, $text)
    $sel.TypeParagraph(); $sel.Style = "List Bullet"; $sel.TypeText($text)
}
function Insert-Table2Col { param($sel, $h1, $h2, $rows)
    $sel.TypeParagraph(); $sel.Style = "Normal"
    $tbl = $sel.Document.Tables.Add($sel.Range, ($rows.Count + 1), 2)
    $tbl.Borders.Enable = $true
    $tbl.Cell(1,1).Range.Text = $h1
    $tbl.Cell(1,2).Range.Text = $h2
    $tbl.Rows(1).Range.Font.Bold = $true
    $tbl.Rows(1).Shading.BackgroundPatternColor = 14935011
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $tbl.Cell($i+2, 1).Range.Text = $rows[$i][0]
        $tbl.Cell($i+2, 2).Range.Text = $rows[$i][1]
    }
    $end = $tbl.Range.Duplicate; $end.Collapse(0); $end.Select()
}

try {
    $doc = $word.Documents.Open($docPath, $false, $false)
    $sel = $word.Selection

    # ---------- STEP 1: Find the BODY "13. Conclusion" heading via Style filter ----------
    Write-Host "Finding body Heading 1 '13. Conclusion'..." -ForegroundColor Cyan
    $sel.HomeKey($wdStory) | Out-Null
    $f = $sel.Find
    $f.ClearFormatting()
    $f.Style = $doc.Styles.Item("Heading 1")
    $f.Text = "13. Conclusion"
    $f.Forward = $true
    $f.Wrap = 0
    $f.MatchCase = $true
    if (-not $f.Execute()) {
        Write-Host "  Style-filtered Find failed; trying without style filter, second occurrence..." -ForegroundColor Yellow
        $sel.HomeKey($wdStory) | Out-Null
        $f.ClearFormatting()
        $f.Text = "13. Conclusion"
        $foundCount = 0
        $bodyFound = $false
        while ($f.Execute()) {
            $foundCount++
            if ($foundCount -ge 2) { $bodyFound = $true; break }
        }
        if (-not $bodyFound) { throw "Could not locate body '13. Conclusion'." }
    }

    Write-Host "  Located. Inserting Section 13 BEFORE it..." -ForegroundColor Green

    # ---------- STEP 2: Insert Section 13 with Track Changes ON ----------
    $doc.TrackRevisions = $true
    $sel.StartOf($wdParagraph, $wdMove) | Out-Null

    # Heading 1
    $sel.Style = "Heading 1"
    $sel.TypeText("13. Medical / Recovery Projects Sourcing and Procurement")
    $sel.TypeParagraph()

    # 13.1
    $sel.Style = "Heading 2"; $sel.TypeText("13.1 Objective")
    Type-Body $sel "To establish clear, enforceable controls over the sourcing and procurement of medical, recovery, rehabilitation, and technical wellness items, ensuring that Supply Chain sources only against complete technical specifications and that responsibility for technical accuracy rests with the requesting and technical departments."

    # 13.2
    Type-Heading2 $sel "13.2 Minimum Technical Requirements Before Sourcing"
    Type-Body $sel "For all medical, recovery, rehabilitation, or technical wellness items, the Supply Chain Department shall not proceed with sourcing unless a complete technical specification is provided in writing by the requesting department. The specification shall include, at minimum:"
    Insert-Table2Col $sel "Requirement" "Details Required" @(
        @("Product type", "Exact item description and intended use"),
        @("Model / size", "Required size, capacity, dimensions, or variant"),
        @("Technical specs", "Power, pressure, resistance level, features, accessories"),
        @("Quantity", "Required quantity per item"),
        @("Brand preference", "Preferred brand or approved alternatives, if any"),
        @("Compliance", "Required certificates, approvals, or regulatory requirements"),
        @("Application", "Medical, recovery, rehab, fitness, or wellness use"),
        @("Delivery location", "Project site and expected delivery date")
    )

    # 13.3
    Type-Heading2 $sel "13.3 Technical Ownership"
    Type-Bullet $sel "The requesting department, Sales, Planning, or BDU, shall be responsible for providing all technical requirements for medical and recovery items."
    Type-Bullet $sel "Supply Chain shall be responsible only for sourcing and commercial negotiation based on the approved technical scope received."
    Type-Bullet $sel "Supply Chain shall not be held responsible for technical selection, medical suitability, or specification gaps when the required technical details are not provided in writing."

    # 13.4
    Type-Heading2 $sel "13.4 BDU and Technical Team Involvement"
    Type-Body $sel "For medical, recovery, and rehabilitation projects, BDU or the assigned technical specialist must be engaged before any RFQ is floated to suppliers. Their role includes:"
    Type-Bullet $sel "Reviewing the client requirement and confirming feasibility."
    Type-Bullet $sel "Defining the correct technical specification or validating the one received."
    Type-Bullet $sel "Confirming acceptable brands or equivalent alternatives."
    Type-Bullet $sel "Reviewing supplier technical offers when required."
    Type-Bullet $sel "Supporting Supply Chain in clarifying technical differences between options."

    # 13.5
    Type-Heading2 $sel "13.5 RFQ Timeline"
    Type-Body $sel "Due to the technical nature of medical and recovery items, the minimum sourcing window shall be two (2) calendar weeks from the date Supply Chain receives complete technical specifications. Shorter windows (for example, two to three working days) are not sufficient for proper supplier identification, clarification, negotiation, and proposal preparation."
    Type-Body $sel "Urgent RFQs received with incomplete specifications or a very short deadline may be accepted only as a documented exception, subject to written management approval acknowledging the associated risks (price, lead time, and technical fit)."

    # 13.6
    Type-Heading2 $sel "13.6 Incomplete Specifications - Completeness Gate"
    Type-Body $sel "If the technical specifications are incomplete, unclear, or too general, Supply Chain shall return the request to the requesting department for clarification before any supplier is approached. The following are examples of requests that fail the completeness gate:"
    Insert-Table2Col $sel "Example Request" "Why It Fails the Gate" @(
        @('"Shockwave machine"', "Multiple types, power levels, and price tiers exist"),
        @('"Recovery equipment"', "Too general - no defined product scope"),
        @('"Medical device"', "Requires technical and regulatory clarification"),
        @('"Massage chair"', "Needs model, functions, size, warranty, and usage")
    )

    # 13.7
    Type-Heading2 $sel "13.7 Supplier Quotation Requirements"
    Type-Body $sel "For medical and recovery project quotations, suppliers must provide BOTH of the following:"
    Type-Bullet $sel "Commercial offer: price, payment terms, delivery lead time, warranty, validity, and installation cost where applicable."
    Type-Bullet $sel "Technical offer / proposal: product datasheet, full specifications, model details, compliance certificates, and accessories included."
    Type-Body $sel "Product catalogues alone shall not be considered a complete technical proposal unless they clearly identify the quoted item and its specifications."

    # 13.8
    Type-Heading2 $sel "13.8 Quotation Validation"
    Type-Body $sel "Any quotation lacking proper technical detail shall be treated as incomplete and shall not be submitted as a final offer to the client unless reviewed and accepted by the concerned technical party (BDU or assigned Technical Owner). Supply Chain may request clarification directly from the supplier where the quoted item does not clearly match the project requirement."

    # 13.9
    Type-Heading2 $sel "13.9 Price Comparison and Variant Control"
    Type-Body $sel "Because many medical and recovery items have multiple variants and price tiers, price comparisons shall be made only between technically equivalent models. Comparison without confirmed alignment on specification, size, capacity, or model shall not constitute a valid comparison and shall not be used to justify a final selection."

    # 13.10
    Type-Heading2 $sel "13.10 Responsibility for Technical Approval"
    Type-Body $sel "Before any purchase order for medical or recovery products is finalized, the final technical approval must be received in writing from the requesting department, BDU, or assigned Technical Owner. Supply Chain approval shall cover commercial aspects only - price, payment terms, delivery terms, and supplier reliability."

    # 13.11
    Type-Heading2 $sel "13.11 Exception Handling for Urgent Projects"
    Type-Body $sel "Where an urgent client requirement is received with insufficient time or incomplete details, Supply Chain shall communicate the risk to the requesting department in writing before proceeding. The communication shall state that the quotation is being prepared based on limited information, and that any technical mismatch, missing accessories, or specification changes discovered after submission may affect the price, lead time, and validity of the final offer."

    # 13.12
    Type-Heading2 $sel "13.12 Required Internal Communication"
    Type-Body $sel "All RFQs related to medical, recovery, or rehabilitation products shall be shared with Supply Chain, BDU, Planning, and the relevant Sales/Project owner at the earliest stage. No department may approach suppliers or submit pricing without proper internal alignment for items requiring technical validation. Any nominated supplier shall be brought into direct contact with Supply Chain and Finance as set out in section 5.4.6."

    # 13.13
    Type-Heading2 $sel "13.13 Policy Statement"
    Type-Body $sel "Supply Chain shall support medical and recovery projects by sourcing qualified suppliers, obtaining competitive quotations, and negotiating the best commercial terms. The accuracy of the technical requirement remains the responsibility of the requesting and technical departments. Supply Chain shall not proceed with procurement based on general descriptions or unclear technical requirements unless management has provided written approval acknowledging the associated risks."

    # 13.14 RACI
    Type-Heading2 $sel "13.14 Roles and Responsibilities"
    Insert-Table2Col $sel "Role" "Responsibility" @(
        @("Requesting Dept. (Sales / Planning)", "Define business need, deliver complete technical spec, provide vendor list (if any), confirm urgency level."),
        @("BDU / Technical Owner", "Review and validate technical scope, confirm acceptable brands, review supplier technical offers, give final technical approval."),
        @("Supply Chain", "Run RFQ, negotiate commercial terms, validate completeness gate, escalate exceptions, issue PO after technical approval."),
        @("Finance", "Confirm budget, validate commercial terms, release payment per agreed schedule."),
        @("Management", "Approve documented exceptions (incomplete spec / urgent timeline) in writing.")
    )

    # 13.15 KPIs
    Type-Heading2 $sel "13.15 Medical / Recovery KPIs"
    Type-Bullet $sel "100% of medical RFQs floated only after the completeness gate is cleared."
    Type-Bullet $sel "Minimum 2-week sourcing window honored in 95% of medical RFQs (exceptions logged with management approval)."
    Type-Bullet $sel "Zero POs issued without written technical approval from BDU or the assigned Technical Owner."
    Type-Bullet $sel "100% of medical quotations include both a commercial and a technical offer."

    # Trailing paragraph break
    $sel.TypeParagraph()

    Write-Host "Section 13 inserted (tracked)." -ForegroundColor Green

    # ---------- STEP 3: Renumber "13. Conclusion" to "14. Conclusion" (without tracking) ----------
    Write-Host "Renumbering Conclusion to Section 14 (untracked, cosmetic)..." -ForegroundColor Cyan
    $doc.TrackRevisions = $false
    $f = $doc.Content.Find
    $f.ClearFormatting()
    $f.Text = "13. Conclusion"
    $f.Replacement.ClearFormatting()
    $f.Replacement.Text = "14. Conclusion"
    $f.Forward = $true
    $f.Wrap = 1  # wdFindContinue
    $f.MatchCase = $true
    $renum = $f.Execute($null,$null,$null,$null,$null,$null,$null,$null,$null,"14. Conclusion",2)
    Write-Host "  Renumber result: $renum" -ForegroundColor Yellow

    # ---------- STEP 4: Update TOC ----------
    Write-Host "Updating TOC..." -ForegroundColor Cyan
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
    Write-Host "  TOC updated." -ForegroundColor Green

    # ---------- Save ----------
    $word.ScreenUpdating = $true
    $doc.Save()
    $doc.Close($false)
    Write-Host "Saved." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
}
finally {
    $word.UserName = $origUser
    $word.UserInitials = $origInit
    try { $word.Quit() } catch {}
}
