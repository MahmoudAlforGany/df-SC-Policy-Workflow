# ========================================================================
# Delta Fitness SC Policy Manual V2.0 - Apply Refinements (Selection-based)
# Author: Mahmoud AlforGany
# ========================================================================

$ErrorActionPreference = "Stop"
$docPath = "C:\Users\alforganym\OneDrive\1 - My Learning Path\1 - Long Term Projects\3 - SC Ploicy WorkFlow\2 - SC Policy Draft (Refined v1 - WIP).docx"

# Word constants
$wdStory          = 6
$wdParagraph      = 4
$wdLine           = 5
$wdMove           = 0
$wdExtend         = 1
$wdCollapseEnd    = 0
$wdCollapseStart  = 1
$wdReplaceAll     = 2

Write-Host "Opening Word..." -ForegroundColor Cyan
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$word.ScreenUpdating = $false

$origUser = $word.UserName
$origInit = $word.UserInitials
$word.UserName = "Mahmoud AlforGany"
$word.UserInitials = "MA"

function Goto-AnchorEnd {
    param([object]$sel, [string]$anchor)
    # Move to top of document, then find anchor; place cursor at end of that paragraph
    $sel.HomeKey($wdStory) | Out-Null
    $f = $sel.Find
    $f.ClearFormatting()
    $f.Text = $anchor
    $f.Forward = $true
    $f.Wrap = 0  # wdFindStop
    $f.MatchCase = $false
    $f.MatchWildcards = $false
    if (-not $f.Execute()) { return $false }
    # Selection is now on matched text. Move to END of containing paragraph.
    $sel.EndOf($wdParagraph, $wdMove) | Out-Null
    return $true
}

function Goto-AnchorStart {
    param([object]$sel, [string]$anchor)
    $sel.HomeKey($wdStory) | Out-Null
    $f = $sel.Find
    $f.ClearFormatting()
    $f.Text = $anchor
    $f.Forward = $true
    $f.Wrap = 0
    $f.MatchCase = $false
    if (-not $f.Execute()) { return $false }
    $sel.StartOf($wdParagraph, $wdMove) | Out-Null
    return $true
}

function Type-Heading {
    param([object]$sel, [string]$level, [string]$text)
    $sel.TypeParagraph()
    $sel.Style = $level
    $sel.TypeText($text)
}

function Type-Body {
    param([object]$sel, [string]$text, [bool]$bold = $false)
    $sel.TypeParagraph()
    $sel.Style = "Normal"
    if ($bold) {
        $sel.Font.Bold = $true
        $sel.TypeText($text)
        $sel.Font.Bold = $false
    } else {
        $sel.TypeText($text)
    }
}

function Type-Bullet {
    param([object]$sel, [string]$text)
    $sel.TypeParagraph()
    $sel.Style = "List Bullet"
    $sel.TypeText($text)
}

function Insert-Table2Col {
    param([object]$sel, [string]$h1, [string]$h2, [array]$rows)
    $sel.TypeParagraph()
    $sel.Style = "Normal"
    # Anchor a new table at the current selection
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
    # Move cursor below the table
    $end = $tbl.Range.Duplicate
    $end.Collapse($wdCollapseEnd)
    $end.Select()
}

try {
    $doc = $word.Documents.Open($docPath, $false, $false)
    $doc.TrackRevisions = $true
    $sel = $word.Selection
    Write-Host "Track Changes ON, Author: $($word.UserName)" -ForegroundColor Green

    # ========================================================================
    # INSERTION 1: 5.4.6 Vendor Communication Protocol
    # Anchor: heading "5.5 Payment Terms" - insert BEFORE it
    # Note: TOC also contains this text; first match is TOC, second is body.
    # Strategy: find twice (skip TOC).
    # ========================================================================
    Write-Host "`n[1/3] Inserting 5.4.6 Vendor Communication Protocol..." -ForegroundColor Cyan

    $sel.HomeKey($wdStory) | Out-Null
    $f = $sel.Find
    $f.ClearFormatting()
    $f.Text = "5.5 Payment Terms"
    $f.Forward = $true
    $f.Wrap = 0
    $f.MatchCase = $true
    $foundCount = 0
    $success = $false
    # First match is likely the TOC entry; we want the actual heading (the body one).
    while ($f.Execute()) {
        $foundCount++
        if ($foundCount -ge 2) { $success = $true; break }
    }
    if (-not $success) {
        # Only one occurrence found - use it
        $sel.HomeKey($wdStory) | Out-Null
        if ($f.Execute()) { $success = $true }
    }

    if ($success) {
        $sel.StartOf($wdParagraph, $wdMove) | Out-Null
        # Insert content BEFORE the section 5.5 heading
        # Type a new paragraph above by typing content + paragraph
        $sel.Style = "Heading 3"
        $sel.TypeText("5.4.6 Vendor Communication Protocol")
        $sel.TypeParagraph()
        $sel.Style = "Normal"
        $sel.TypeText("To ensure commercial integrity, consistent terms, and a single audit trail, all communication with vendors and service providers shall follow the protocol below.")

        Type-Body $sel "Communication Ownership" $true

        $bullets1 = @(
            "All negotiations, pricing discussions, and commercial communications with vendors must be conducted by, or under the direct supervision of, the Procurement / Supply Chain team.",
            "The Sales team shall not act as the sole focal point with suppliers. Sales may participate in technical or scoping discussions, but commercial negotiation authority remains with Supply Chain.",
            "Where a supplier or service provider is nominated by another department (Sales, BDU, Technical), that supplier must be brought into direct communication with the relevant support functions, Supply Chain and Finance, and not routed indirectly through the nominating department.",
            "All vendor lists, including nominated or preferred suppliers, must be submitted to Supply Chain in writing before the RFQ stage, so that direct commercial communication channels can be established.",
            "Any exception to this protocol requires written approval from the Supply Chain Manager."
        )
        foreach ($b in $bullets1) { Type-Bullet $sel $b }

        Type-Body $sel "Rationale" $true
        Type-Body $sel "Centralizing vendor communication through Supply Chain prevents fragmented pricing, conflicting commitments, and loss of negotiation leverage. It also protects the company from accountability gaps where commercial terms have been agreed informally outside the procurement channel."

        # Final paragraph break before the heading we anchored on
        $sel.TypeParagraph()

        Write-Host "  5.4.6 inserted." -ForegroundColor Green
    } else {
        Write-Host "  [WARN] Could not anchor 5.4.6 insertion." -ForegroundColor Red
    }

    # ========================================================================
    # INSERTION 2: 6.3.2.A Project Transportation & Wared Coordination
    # Anchor: heading "6.3.3 Customs Clearance Process" - insert BEFORE it
    # ========================================================================
    Write-Host "`n[2/3] Inserting 6.3.2.A Project Transportation..." -ForegroundColor Cyan

    $sel.HomeKey($wdStory) | Out-Null
    $f = $sel.Find
    $f.ClearFormatting()
    $f.Text = "6.3.3 Customs Clearance Process"
    $f.Forward = $true
    $f.Wrap = 0
    $f.MatchCase = $true
    $foundCount = 0
    $success2 = $false
    while ($f.Execute()) {
        $foundCount++
        if ($foundCount -ge 2) { $success2 = $true; break }
    }
    if (-not $success2) {
        $sel.HomeKey($wdStory) | Out-Null
        if ($f.Execute()) { $success2 = $true }
    }

    if ($success2) {
        $sel.StartOf($wdParagraph, $wdMove) | Out-Null

        $sel.Style = "Heading 3"
        $sel.TypeText("6.3.2.A Project Transportation and Wared Coordination")
        $sel.TypeParagraph()
        $sel.Style = "Normal"
        $sel.TypeText("This sub-section governs the movement of project-bound goods between Delta Fitness internal hubs and customer sites, with specific controls for Riyadh, Khobar, and the Wared transit location.")

        Type-Body $sel "Shipment Pre-Approval (Riyadh and Khobar)" $true

        $b2a = @(
            "Any shipment planned for a Riyadh or Khobar project site must be pre-approved in writing by the Technical Support (TS) team prior to dispatch.",
            "TS approval shall be granted only after the expected delivery schedule has been confirmed against the customer's site readiness, ensuring goods are not dispatched ahead of installation capacity.",
            "Approval evidence (email confirmation or signed dispatch sheet) must be attached to the shipment record before the warehouse releases goods."
        )
        foreach ($b in $b2a) { Type-Bullet $sel $b }

        Type-Body $sel "Wared Dwell-Time SLA" $true

        $b2b = @(
            "Standard dwell limit: No project shipment may remain at the Wared transit location for more than 1 calendar week (7 days) from arrival.",
            "Maximum extension: In special cases, dwell time may be extended to a maximum of 2 calendar weeks (14 days), subject to written justification approved by the TS Manager and Supply Chain Manager.",
            "Mandatory return: After 14 days, the items shall be returned to the originating warehouse location. The return movement is non-discretionary.",
            "Escalation email: Concurrent with the return, an escalation email shall be issued to the SC Manager, TS Manager, and the responsible Sales / Project owner, stating the reason the shipment exceeded the dwell limit, the customer-side cause, and the corrective action.",
            "Weekly review: Logistics and TS shall jointly review the Wared dwell tracker every Monday and flag any shipment approaching the 7-day limit."
        )
        foreach ($b in $b2b) { Type-Bullet $sel $b }

        Type-Body $sel "KPI" $true
        Type-Bullet $sel "Zero shipments exceeding the 14-day Wared dwell limit without documented escalation in a rolling quarter."

        $sel.TypeParagraph()

        Write-Host "  6.3.2.A inserted." -ForegroundColor Green
    } else {
        Write-Host "  [WARN] Could not anchor 6.3.2.A insertion." -ForegroundColor Red
    }

    # ========================================================================
    # INSERTION 3: NEW SECTION 13 (Medical/Recovery) + renumber Conclusion
    # ========================================================================
    Write-Host "`n[3/3] Inserting Section 13 (Medical/Recovery)..." -ForegroundColor Cyan

    # Renumber: "13. Conclusion" -> "14. Conclusion" (both TOC and body)
    $sel.HomeKey($wdStory) | Out-Null
    $f = $sel.Find
    $f.ClearFormatting()
    $f.Text = "13. Conclusion"
    $f.Replacement.ClearFormatting()
    $f.Replacement.Text = "14. Conclusion"
    $f.Forward = $true
    $f.Wrap = 0
    $renum = $f.Execute($null,$null,$null,$null,$null,$null,$null,$null,$null,"14. Conclusion",$wdReplaceAll)
    if ($renum) { Write-Host "  Renumbered Conclusion to Section 14." -ForegroundColor Green }
    else { Write-Host "  [WARN] Could not renumber Conclusion." -ForegroundColor Yellow }

    # Insert Section 13 BEFORE "14. Conclusion" body heading (skip TOC entry)
    $sel.HomeKey($wdStory) | Out-Null
    $f = $sel.Find
    $f.ClearFormatting()
    $f.Text = "14. Conclusion"
    $f.Forward = $true
    $f.Wrap = 0
    $f.MatchCase = $true
    $foundCount = 0
    $success3 = $false
    while ($f.Execute()) {
        $foundCount++
        if ($foundCount -ge 2) { $success3 = $true; break }
    }
    if (-not $success3) {
        $sel.HomeKey($wdStory) | Out-Null
        if ($f.Execute()) { $success3 = $true }
    }

    if ($success3) {
        $sel.StartOf($wdParagraph, $wdMove) | Out-Null

        # Heading 1
        $sel.Style = "Heading 1"
        $sel.TypeText("13. Medical / Recovery Projects Sourcing and Procurement")
        $sel.TypeParagraph()

        # 13.1
        $sel.Style = "Heading 2"
        $sel.TypeText("13.1 Objective")
        Type-Body $sel "To establish clear, enforceable controls over the sourcing and procurement of medical, recovery, rehabilitation, and technical wellness items, ensuring that Supply Chain sources only against complete technical specifications and that responsibility for technical accuracy rests with the requesting and technical departments."

        # 13.2
        Type-Heading $sel "Heading 2" "13.2 Minimum Technical Requirements Before Sourcing"
        Type-Body $sel "For all medical, recovery, rehabilitation, or technical wellness items, the Supply Chain Department shall not proceed with sourcing unless a complete technical specification is provided in writing by the requesting department. The specification shall include, at minimum:"

        $reqRows = @(
            @("Product type", "Exact item description and intended use"),
            @("Model / size", "Required size, capacity, dimensions, or variant"),
            @("Technical specs", "Power, pressure, resistance level, features, accessories"),
            @("Quantity", "Required quantity per item"),
            @("Brand preference", "Preferred brand or approved alternatives, if any"),
            @("Compliance", "Required certificates, approvals, or regulatory requirements"),
            @("Application", "Medical, recovery, rehab, fitness, or wellness use"),
            @("Delivery location", "Project site and expected delivery date")
        )
        Insert-Table2Col $sel "Requirement" "Details Required" $reqRows

        # 13.3
        Type-Heading $sel "Heading 2" "13.3 Technical Ownership"
        foreach ($b in @(
            "The requesting department, Sales, Planning, or BDU, shall be responsible for providing all technical requirements for medical and recovery items.",
            "Supply Chain shall be responsible only for sourcing and commercial negotiation based on the approved technical scope received.",
            "Supply Chain shall not be held responsible for technical selection, medical suitability, or specification gaps when the required technical details are not provided in writing."
        )) { Type-Bullet $sel $b }

        # 13.4
        Type-Heading $sel "Heading 2" "13.4 BDU and Technical Team Involvement"
        Type-Body $sel "For medical, recovery, and rehabilitation projects, BDU or the assigned technical specialist must be engaged before any RFQ is floated to suppliers. Their role includes:"
        foreach ($b in @(
            "Reviewing the client requirement and confirming feasibility.",
            "Defining the correct technical specification or validating the one received.",
            "Confirming acceptable brands or equivalent alternatives.",
            "Reviewing supplier technical offers when required.",
            "Supporting Supply Chain in clarifying technical differences between options."
        )) { Type-Bullet $sel $b }

        # 13.5
        Type-Heading $sel "Heading 2" "13.5 RFQ Timeline"
        Type-Body $sel "Due to the technical nature of medical and recovery items, the minimum sourcing window shall be two (2) calendar weeks from the date Supply Chain receives complete technical specifications. Shorter windows (for example, two to three working days) are not sufficient for proper supplier identification, clarification, negotiation, and proposal preparation."
        Type-Body $sel "Urgent RFQs received with incomplete specifications or a very short deadline may be accepted only as a documented exception, subject to written management approval acknowledging the associated risks (price, lead time, and technical fit)."

        # 13.6
        Type-Heading $sel "Heading 2" "13.6 Incomplete Specifications - Completeness Gate"
        Type-Body $sel "If the technical specifications are incomplete, unclear, or too general, Supply Chain shall return the request to the requesting department for clarification before any supplier is approached. The following are examples of requests that fail the completeness gate:"

        $incRows = @(
            @('"Shockwave machine"', "Multiple types, power levels, and price tiers exist"),
            @('"Recovery equipment"', "Too general - no defined product scope"),
            @('"Medical device"',     "Requires technical and regulatory clarification"),
            @('"Massage chair"',      "Needs model, functions, size, warranty, and usage")
        )
        Insert-Table2Col $sel "Example Request" "Why It Fails the Gate" $incRows

        # 13.7
        Type-Heading $sel "Heading 2" "13.7 Supplier Quotation Requirements"
        Type-Body $sel "For medical and recovery project quotations, suppliers must provide BOTH of the following:"
        foreach ($b in @(
            "Commercial offer: price, payment terms, delivery lead time, warranty, validity, and installation cost where applicable.",
            "Technical offer / proposal: product datasheet, full specifications, model details, compliance certificates, and accessories included."
        )) { Type-Bullet $sel $b }
        Type-Body $sel "Product catalogues alone shall not be considered a complete technical proposal unless they clearly identify the quoted item and its specifications."

        # 13.8
        Type-Heading $sel "Heading 2" "13.8 Quotation Validation"
        Type-Body $sel "Any quotation lacking proper technical detail shall be treated as incomplete and shall not be submitted as a final offer to the client unless reviewed and accepted by the concerned technical party (BDU or assigned Technical Owner). Supply Chain may request clarification directly from the supplier where the quoted item does not clearly match the project requirement."

        # 13.9
        Type-Heading $sel "Heading 2" "13.9 Price Comparison and Variant Control"
        Type-Body $sel "Because many medical and recovery items have multiple variants and price tiers, price comparisons shall be made only between technically equivalent models. Comparison without confirmed alignment on specification, size, capacity, or model shall not constitute a valid comparison and shall not be used to justify a final selection."

        # 13.10
        Type-Heading $sel "Heading 2" "13.10 Responsibility for Technical Approval"
        Type-Body $sel "Before any purchase order for medical or recovery products is finalized, the final technical approval must be received in writing from the requesting department, BDU, or assigned Technical Owner. Supply Chain approval shall cover commercial aspects only - price, payment terms, delivery terms, and supplier reliability."

        # 13.11
        Type-Heading $sel "Heading 2" "13.11 Exception Handling for Urgent Projects"
        Type-Body $sel "Where an urgent client requirement is received with insufficient time or incomplete details, Supply Chain shall communicate the risk to the requesting department in writing before proceeding. The communication shall state that the quotation is being prepared based on limited information, and that any technical mismatch, missing accessories, or specification changes discovered after submission may affect the price, lead time, and validity of the final offer."

        # 13.12
        Type-Heading $sel "Heading 2" "13.12 Required Internal Communication"
        Type-Body $sel "All RFQs related to medical, recovery, or rehabilitation products shall be shared with Supply Chain, BDU, Planning, and the relevant Sales/Project owner at the earliest stage. No department may approach suppliers or submit pricing without proper internal alignment for items requiring technical validation. Any nominated supplier shall be brought into direct contact with Supply Chain and Finance as set out in section 5.4.6."

        # 13.13
        Type-Heading $sel "Heading 2" "13.13 Policy Statement"
        Type-Body $sel "Supply Chain shall support medical and recovery projects by sourcing qualified suppliers, obtaining competitive quotations, and negotiating the best commercial terms. The accuracy of the technical requirement remains the responsibility of the requesting and technical departments. Supply Chain shall not proceed with procurement based on general descriptions or unclear technical requirements unless management has provided written approval acknowledging the associated risks."

        # 13.14 RACI
        Type-Heading $sel "Heading 2" "13.14 Roles and Responsibilities"
        $raciRows = @(
            @("Requesting Dept. (Sales / Planning)", "Define business need, deliver complete technical spec, provide vendor list (if any), confirm urgency level."),
            @("BDU / Technical Owner", "Review and validate technical scope, confirm acceptable brands, review supplier technical offers, give final technical approval."),
            @("Supply Chain", "Run RFQ, negotiate commercial terms, validate completeness gate, escalate exceptions, issue PO after technical approval."),
            @("Finance", "Confirm budget, validate commercial terms, release payment per agreed schedule."),
            @("Management", "Approve documented exceptions (incomplete spec / urgent timeline) in writing.")
        )
        Insert-Table2Col $sel "Role" "Responsibility" $raciRows

        # 13.15 KPIs
        Type-Heading $sel "Heading 2" "13.15 Medical / Recovery KPIs"
        foreach ($k in @(
            "100% of medical RFQs floated only after the completeness gate is cleared.",
            "Minimum 2-week sourcing window honored in 95% of medical RFQs (exceptions logged with management approval).",
            "Zero POs issued without written technical approval from BDU or the assigned Technical Owner.",
            "100% of medical quotations include both a commercial and a technical offer."
        )) { Type-Bullet $sel $k }

        # Final paragraph break before the section 14 Conclusion heading
        $sel.TypeParagraph()

        Write-Host "  Section 13 inserted." -ForegroundColor Green
    } else {
        Write-Host "  [WARN] Could not anchor Section 13 insertion." -ForegroundColor Red
    }

    # ========================================================================
    # Update TOC
    # ========================================================================
    Write-Host "`nUpdating TOC..." -ForegroundColor Cyan
    try {
        foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
        Write-Host "  TOC updated." -ForegroundColor Green
    } catch {
        Write-Host "  [WARN] TOC update issue: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    $word.ScreenUpdating = $true
    $doc.Save()
    $doc.Close($false)
    Write-Host "`nSaved: $docPath" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "At line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
}
finally {
    $word.UserName = $origUser
    $word.UserInitials = $origInit
    try { $word.Quit() } catch {}
    Write-Host "Word closed." -ForegroundColor Gray
}
