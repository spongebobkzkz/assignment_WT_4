$ErrorActionPreference = "Stop"

$tempDir = "C:\Users\spong\temp_report"
if (-not (Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
}

$sourceScreenshots = Join-Path $PSScriptRoot "screenshots"
Copy-Item "$sourceScreenshots\*.png" -Destination $tempDir -Force

$docxTemp = "$tempDir\Assignment4_Beisenbay_Estai.docx"
$pdfTemp = "$tempDir\Assignment4_Beisenbay_Estai.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Add()
$selection = $word.Selection

$wdAlignLeft = 0
$wdAlignCenter = 1
$wdAlignRight = 2

$colTeal = 6253327        # #0F6B5F
$colDarkGreen = 2764822   # #16302B
$colGray = 8421504
$colBlue = 13382451       # #3366BB
$colHeaderGray = 14869218 # light gray

function Add-Heading {
    param([string]$headingText, [int]$level)
    $selection.Font.Name = "Montserrat"
    $selection.Font.Bold = $true
    if ($level -eq 1) {
        $selection.Font.Size = 17
        $selection.Font.Color = $colTeal
        $selection.ParagraphFormat.SpaceBefore = 14
        $selection.ParagraphFormat.SpaceAfter = 6
    } elseif ($level -eq 2) {
        $selection.Font.Size = 13
        $selection.Font.Color = $colDarkGreen
        $selection.ParagraphFormat.SpaceBefore = 10
        $selection.ParagraphFormat.SpaceAfter = 4
    } else {
        $selection.Font.Size = 11
        $selection.Font.Color = 0
        $selection.ParagraphFormat.SpaceBefore = 6
        $selection.ParagraphFormat.SpaceAfter = 2
    }
    $selection.TypeText($headingText)
    $selection.TypeParagraph()
}

function Add-Text {
    param([string]$paragraphText)
    $selection.Font.Name = "Calibri"
    $selection.Font.Bold = $false
    $selection.Font.Size = 11
    $selection.Font.Color = 0
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 5
    $selection.ParagraphFormat.LineSpacing = 13.5
    $selection.TypeText($paragraphText)
    $selection.TypeParagraph()
}

function Add-Image {
    param([string]$imagePath, [float]$imgWidth)
    if (Test-Path $imagePath) {
        $shape = $selection.InlineShapes.AddPicture($imagePath)
        $shape.Width = $imgWidth
        $selection.ParagraphFormat.Alignment = $wdAlignCenter
        $selection.TypeParagraph()
        $selection.ParagraphFormat.Alignment = $wdAlignLeft
    }
}

# ================= TITLE & HEADER =================
$selection.Font.Name = "Montserrat"
$selection.Font.Size = 22
$selection.Font.Bold = $true
$selection.Font.Color = $colTeal
$selection.ParagraphFormat.Alignment = $wdAlignCenter
$selection.ParagraphFormat.SpaceAfter = 6
$selection.TypeText("Assignment #4. Forms and Validation")
$selection.TypeParagraph()

$selection.Font.Size = 13
$selection.Font.Bold = $false
$selection.Font.Color = $colGray
$selection.TypeText("HTML5 Native Validation + Bootstrap 5.3 Components")
$selection.TypeParagraph()
$selection.TypeParagraph()

# Student metadata block
$selection.ParagraphFormat.Alignment = $wdAlignLeft
$selection.Font.Name = "Montserrat"
$selection.Font.Size = 11
$selection.Font.Bold = $true
$selection.Font.Color = 0
$selection.TypeText("Name: ")
$selection.Font.Bold = $false
$selection.TypeText("Beisenbay Estai`n")

$selection.Font.Bold = $true
$selection.TypeText("Group: ")
$selection.Font.Bold = $false
$selection.TypeText("IT-2511`n")

$selection.Font.Bold = $true
$selection.TypeText("Course: ")
$selection.Font.Bold = $false
$selection.TypeText("Web Technologies (2nd Year, 1st Trimester)`n")

$selection.Font.Bold = $true
$selection.TypeText("University: ")
$selection.Font.Bold = $false
$selection.TypeText("Astana IT University`n`n")

# ================= INTRODUCTION =================
Add-Heading "Introduction" 1
Add-Text "In this assignment, I designed and implemented three web form applications focusing on Bootstrap form components, native HTML5 validation constraints, usability, and accessibility standards. The goals of this practical work are:"
Add-Text "1. Build responsive forms using Bootstrap 5.3 form components and grid system (Task 1).`n2. Validate form data with native HTML attributes only, showing clear feedback messages without any JavaScript or novalidate attributes (Task 2).`n3. Audit, document, and remediate usability, accessibility, and syntax issues in a flawed form, ensuring full W3C validity and keyboard navigation (Task 3)."

# ================= TASK 1 =================
Add-Heading "Task 1 - Book a Train Ticket (Bootstrap Forms + Custom CSS)" 1
Add-Text "In Task 1, I reproduced the train ticket booking interface according to the provided specification. The form was built using standard Bootstrap 5.3 components and restyled using custom CSS with the exact palette:"
Add-Text "- Primary Teal: #0F6B5F`n- Accent Orange/Gold: #F2A93B`n- Dark Forest Text: #16302B`n- Background Page: #F4F7F5`n- Light Border: #E6EFEC`n- Typography: Google Font Montserrat"

Add-Text "Key Technical Highlights of Task 1:"
Add-Text "- Radio Toggle for Trip Type: Created a segmented pill toggle using two native radio inputs (input type='radio' class='btn-check' name='tripType') styled as rounded buttons.`n- Floating Labels: Used Bootstrap .form-floating for 'From' (Astana) and 'To' (Almaty).`n- Three-Column Layout: Departure date, Return date, and Passengers select (1-6 adults) are placed in a responsive row using col-12 col-md-4.`n- Carriage Class: Inline radio inputs for Seat, Coupe, and Sleeper.`n- Input Group: The telephone input contains a custom styled teal prefix (+7) using Bootstrap .input-group and .input-group-text.`n- Checkboxes and Switches: Built the SMS notification switch (.form-switch) and carriage rules checkbox.`n- Responsive Button Order: Crucially, on a phone, 'Find trains' is positioned ABOVE 'Clear' without altering the HTML order. This is achieved natively using Bootstrap flex utility classes: .d-flex .flex-column-reverse .flex-md-row .justify-content-md-end. The 'Clear' button uses type='reset' to empty the form cleanly without JavaScript."

Add-Heading "Screenshots of Task 1:" 2
Add-Text "1) Laptop View (1280px screen width):"
Add-Image "$tempDir\task1_1280px.png" 460

Add-Text "`n2) Phone View (375px screen width):"
Add-Image "$tempDir\task1_375px.png" 210

# ================= TASK 2 =================
Add-Heading "Task 2 - Registration for AITU Hack 2026 (HTML Validation)" 1
Add-Text "In Task 2, I implemented a student hackathon registration form that strictly relies on native HTML5 validation attributes without using the pattern attribute, without any JavaScript, and without novalidate. Browser validation handles the data checks natively."

Add-Text "Implementation of :user-invalid and :user-valid:"
Add-Text "A key requirement is that validation error indicators (red borders and .invalid-feedback messages) must appear ONLY after the user interacts with or modifies a field, not when the page first loads. Bootstrap hides .invalid-feedback by default. Using modern CSS pseudo-classes:"
Add-Text "input:user-invalid ~ .invalid-feedback, select:user-invalid ~ .invalid-feedback, textarea:user-invalid ~ .invalid-feedback { display: block; } input:user-invalid { border-color: #dc3545; } input:user-valid { border-color: #198754; }"
Add-Text "This ensures a pristine initial experience where empty required fields do not immediately flash red upon opening."

Add-Heading "Task 2 Results and Attribute Verification Table:" 2

$tableData = @(
    @("Field", "Rule", "HTML Attributes Used", "Must Accept", "Must Reject", "Result"),
    @("Full name *", "2 to 40 characters", "type='text' minlength='2' maxlength='40' required", "Anna-Maria Lee", "A", "Pass"),
    @("University email *", "a valid email address", "type='email' required", "a.fedenko@astanait.edu.kz", "a.fedenko, a.fedenko@", "Pass"),
    @("Age *", "whole number 16 to 30", "type='number' min='16' max='30' step='1' required", "18", "15, 31", "Pass"),
    @("Team size *", "whole number 1 to 4", "type='number' min='1' max='4' step='1' required", "4", "5, 2.5", "Pass"),
    @("Track *", "Web, Mobile, AI, GameDev", "select required + disabled hidden default option", "Web", "Choose...", "Pass"),
    @("Arrival day *", "20 to 22 Nov 2026", "type='date' min='2026-11-20' max='2026-11-22' required", "21.11.2026", "25.11.2026", "Pass"),
    @("Check-in time *", "09:00 to 18:00, step 30 min", "type='time' min='09:00' max='18:00' step='1800' required", "10:30", "10:15, 19:00", "Pass"),
    @("Entry fee *", "0 to 5000 KZT, step 500 KZT", "type='number' min='0' max='5000' step='500' required", "1500", "1250, 6000", "Pass"),
    @("Experience level *", "one of 3 radio buttons", "type='radio' name='experienceLevel' required", "Beginner", "nothing chosen", "Pass"),
    @("GitHub profile", "optional, a full link", "type='url' (no required attribute)", "https://github.com/...", "github.com/...", "Pass"),
    @("Password *", "8 to 20 characters", "type='password' minlength='8' maxlength='20' required", "Astana2026", "short", "Pass"),
    @("Motivation *", "50 to 300 characters", "textarea minlength='50' maxlength='300' required", "120 characters", "49 characters", "Pass"),
    @("Rules agreement *", "checkbox", "type='checkbox' required", "ticked", "not ticked", "Pass")
)

$t2Rows = $tableData.Count
$t2Cols = $tableData[0].Count
$table2 = $doc.Tables.Add($selection.Range, $t2Rows, $t2Cols)
$table2.Borders.Enable = $true
for ($r = 0; $r -lt $t2Rows; $r++) {
    for ($c = 0; $c -lt $t2Cols; $c++) {
        $cell = $table2.Cell($r + 1, $c + 1)
        $cell.Range.Text = $tableData[$r][$c]
        $cell.Range.Font.Name = "Calibri"
        $cell.Range.Font.Size = 9.5
        if ($r -eq 0) {
            $cell.Range.Font.Bold = $true
            $cell.Shading.BackgroundPatternColor = $colHeaderGray
        }
    }
}
$selection.SetRange($table2.Range.End, $table2.Range.End)
$selection.TypeParagraph()

Add-Text "`nScreenshot of Task 2 Form and Live Validation:"
Add-Image "$tempDir\task2_1280px.png" 460

# ================= TASK 3 =================
Add-Heading "Task 3 - Fix a Bad Form (HTML + Bootstrap)" 1
Add-Text "A cafe reservation form was provided with multiple severe usability, accessibility, and markup validity flaws. Below is the comprehensive audit identifying 12 critical defects, explaining why each is problematic for users, and detailing how it was resolved in task3.html."

Add-Heading "Table of Identified Problems and Fixes (with 2 Alternatives):" 2

$probData = @(
    @("#", "Line of Code", "Why It Is a Problem for the User", "How We Fixed It and Alternatives"),
    @("1", "<input type='text' placeholder='Your name'>", "Missing label and name attribute. Placeholder vanishes on typing, harming short-term memory. Screen readers cannot announce the field name. Value is lost on submit.", "Added <label for='resName'>Your Name</label> and id='resName' name='customerName' required autocomplete='name'."),
    @("2", "<input type='text' placeholder='Email'>", "Wrong input type and missing label. Mobile keyboards do not show email shortcuts (@, .com) and browser skips email syntax validation.", "Changed to type='email', added <label for='resEmail'>, id='resEmail' name='customerEmail' required."),
    @("3", "<label>Phone</label>`n<input type='text' id='phone'>", "Label not connected via 'for' attribute; clicking 'Phone' does not focus the input. type='text' does not trigger numeric keypad.", "CHOSEN FIX: Added for='resPhone' to label, changed to type='tel' name='customerPhone' autocomplete='tel'.`n`nALTERNATIVE FIX: Wrap input inside label: <label>Phone <input type='tel'></label>.`n`nWHY WE CHOSE OURS: Bootstrap 5 styles and grid layouts work best with separate .form-label and .form-control elements without layout glitches."),
    @("4", "<label for='guests'>Guests</label>`n<input type='text' id='people'>", "ID mismatch! Label points to 'guests' but input is 'people'. Clicking label fails. type='text' allows non-numbers.", "Synchronized IDs: for='resGuests' and id='resGuests'. Changed to type='number' min='1' max='30' step='1' name='guestsCount'."),
    @("5", "<label for='date'>Date</label>`n<input type='text' id='date' placeholder='dd/mm/yyyy'>", "Manual date typing leads to format errors. No datepicker. Also id='date' is duplicated in the radio button below!", "Changed to type='date', assigned unique id='resDate' with connected label, and added name='reservationDate'."),
    @("6", "<input type='radio' name='inside' id='date'> Inside`n<input type='radio' name='terrace'> Terrace", "Different names prevent radio mutual exclusion (both can be checked). Duplicate id='date' violates W3C. Missing label tags with 'for' attribute.", "Unified name='seatingPreference' for both radios, assigned unique IDs (seatingInside, seatingTerrace), added <label class='form-check-label'> and explicit values."),
    @("7", "<p>Seating:</p>", "Unsemantic paragraph gives no programmatic association to the radio options for screen readers.", "Replaced with semantic <fieldset><legend class='form-label'>Seating Preference</legend>...</fieldset>."),
    @("8", "<select><option>Lunch</option><option>Dinner</option></select>", "Unlabeled dropdown fails accessibility (WCAG 2.1). Missing name attribute causes data loss on submission.", "Added <label for='resMealSlot'>Meal Slot</label>, id='resMealSlot' name='mealSlot' required, and explicit option values."),
    @("9", "<textarea placeholder='Comments'></textarea>", "Missing label, id, and name attribute. Text cannot be sent to server and cannot be identified by screen readers.", "Added <label for='resComments'>Comments and Requests</label>, id='resComments' name='customerComments' rows='3'."),
    @("10", "<input type='checkbox'> I agree to the terms", "Unlinked label. Mobile users cannot tap the label text to check the box. Missing name and required attributes.", "Wrapped in Bootstrap .form-check with <input type='checkbox' class='form-check-input' id='resTerms' required> and <label for='resTerms'>."),
    @("11", "<div class='btn btn-primary'>Reserve</div>", "CRITICAL FLAW: A <div> cannot be focused via keyboard Tab key, cannot be activated with Enter/Space, and NEVER submits the form!", "CHOSEN FIX: Replaced with real HTML <button type='submit' class='btn btn-primary'>Reserve Table</button>.`n`nALTERNATIVE FIX: Use <input type='submit' value='Reserve Table' class='btn btn-primary'>.`n`nWHY WE CHOSE OURS: The <button> element is more flexible, supports child HTML elements/icons, and is Bootstrap 5 recommended standard."),
    @("12", "<form>", "Missing action and method attributes. Defaults to GET and exposes sensitive customer parameters in the URL query string.", "Added action='#' method='post'.")
)

$t3Rows = $probData.Count
$t3Cols = $probData[0].Count
$table3 = $doc.Tables.Add($selection.Range, $t3Rows, $t3Cols)
$table3.Borders.Enable = $true
for ($r = 0; $r -lt $t3Rows; $r++) {
    for ($c = 0; $c -lt $t3Cols; $c++) {
        $cell = $table3.Cell($r + 1, $c + 1)
        $cell.Range.Text = $probData[$r][$c]
        $cell.Range.Font.Name = "Calibri"
        $cell.Range.Font.Size = 9
        if ($r -eq 0) {
            $cell.Range.Font.Bold = $true
            $cell.Shading.BackgroundPatternColor = $colHeaderGray
        }
    }
}
$selection.SetRange($table3.Range.End, $table3.Range.End)
$selection.TypeParagraph()

Add-Text "`nScreenshot of Task 3 Rebuilt Form and Comparison Table:"
Add-Image "$tempDir\task3_1280px.png" 460

# ================= CONCLUSION =================
Add-Heading "Conclusion" 1
Add-Text "Through completing Assignment #4, I acquired essential practical skills in designing modern, responsive, and accessible web forms:"
Add-Text "1. Mastering Bootstrap Form Controls: Learned how to combine native HTML elements with Bootstrap grid, input groups, floating labels, segmented buttons, switches, and responsive flex utilities.`n2. Pure HTML5 Form Validation: Gained deep understanding of how to enforce strict business rules (ranges, steps, length limits, time intervals) purely through native HTML attributes without a single line of JavaScript.`n3. Modern CSS Validation States: Applied the :user-invalid and :user-valid pseudo-classes so that error indicators appear gracefully only after user interaction.`n4. Accessibility and Usability (WCAG): Understood the critical importance of connected labels (<label for='...'>), semantic <fieldset> and <legend> structures, correct input types for virtual mobile keyboards, and using genuine <button type='submit'> elements for keyboard navigability.`n5. Responsive Form Design: Successfully handled mobile-first layout adaptations, such as reversing button order without modifying HTML structure."

# ================= GITHUB REPOSITORY =================
Add-Heading "GitHub Repository URL" 1
Add-Text "All source code, stylesheets, screenshots, and documentation for Assignment #4 are published at:"
$selection.Font.Name = "Consolas"
$selection.Font.Size = 10.5
$selection.Font.Bold = $true
$selection.Font.Color = $colBlue
$selection.TypeText("Source Code Repository: https://github.com/spongebobkzkz/assignment_WT_4`n")
$selection.TypeText("Live Website Deployment: https://spongebobkzkz.github.io/assignment_WT_4/`n")

# Save as DOCX
$doc.SaveAs2($docxTemp, 16)

# Save as PDF using ExportAsFixedFormat
$doc.ExportAsFixedFormat($pdfTemp, 17)

$doc.Close(0)
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

$finalTargetDir = $PSScriptRoot
Copy-Item $docxTemp (Join-Path $finalTargetDir "Assignment4_Beisenbay_Estai.docx") -Force
Copy-Item $pdfTemp (Join-Path $finalTargetDir "Assignment4_Beisenbay_Estai.pdf") -Force

Write-Host "SUCCESS! Generated and copied to:"
Write-Host "  $(Join-Path $finalTargetDir 'Assignment4_Beisenbay_Estai.docx')"
Write-Host "  $(Join-Path $finalTargetDir 'Assignment4_Beisenbay_Estai.pdf')"
