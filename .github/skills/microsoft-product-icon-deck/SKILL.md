---
name: microsoft-product-icon-deck
description: "Use when: creating or updating a PowerPoint deck of Microsoft product icons by category, including M365, D365, Azure, Power Platform, and Entra icons. Helps PPT Copilot generate a categorized Microsoft icon deck from public Microsoft icon sources."
---

# Microsoft Product Icon Deck Skill

Use this skill when the user asks for a PowerPoint icon deck, Microsoft product icon catalog, or a categorized deck for Microsoft product families such as M365, D365, Azure, Power Platform, or Entra.

## Output

Generate or update `Microsoft_Product_Icons_By_Category_Public.pptx`.

## Workflow

1. Use the repository script `scripts/generate-microsoft-icons-deck.js`.
2. Build the public deck with:

   ```powershell
   $env:PUBLIC_RELEASE='1'
   $env:OUTPUT_DECK='Microsoft_Product_Icons_By_Category_Public.pptx'
   npm run build
   Remove-Item Env:PUBLIC_RELEASE,Env:OUTPUT_DECK
   ```

3. Validate the generated deck content:

   ```powershell
   python -m markitdown Microsoft_Product_Icons_By_Category_Public.pptx | Select-String -Pattern 'xxxx','lorem','ipsum','placeholder','Missing icons','not found','Generation note' -CaseSensitive:$false
   ```

4. If PowerPoint is available, export slides to PNG for visual QA:

   ```powershell
   $pptPath = Join-Path (Get-Location) 'Microsoft_Product_Icons_By_Category_Public.pptx'
   $outDir = Join-Path (Get-Location) 'qa-slides'
   New-Item -ItemType Directory -Force -Path $outDir | Out-Null
   Remove-Item $outDir\*.png -Force -ErrorAction SilentlyContinue
   $powerPoint = New-Object -ComObject PowerPoint.Application
   $presentation = $powerPoint.Presentations.Open($pptPath, $true, $false, $false)
   foreach ($slide in $presentation.Slides) { $slide.Export((Join-Path $outDir ('slide-' + ('{0:d2}' -f $slide.SlideIndex) + '.png')), 'PNG', 1600, 900) }
   $presentation.Close()
   $powerPoint.Quit()
   ```

## Public Release Rule

For public GitHub publishing, always set `PUBLIC_RELEASE=1`. Do not publish private roadmap decks, extracted roadmap media, `official-icon-assets/`, `generated-assets/`, `qa-slides/`, or `node_modules/`.

## Icon Terms

Use only under the Microsoft icon terms published on Microsoft Learn. Keep product names close to icons. Do not crop, flip, rotate, distort, or use Microsoft product icons to represent another product or service.
