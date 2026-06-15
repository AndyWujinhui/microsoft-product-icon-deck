---
name: microsoft-product-icon-deck
description: "Use when: generating, updating, or inserting a PowerPoint deck of Microsoft product icons grouped by category, including M365, D365, Azure, Power Platform, and Entra. This skill can reuse the published icon deck or regenerate it from the local project."
---

# Microsoft Product Icon Deck

Use this skill when a user asks PPT Copilot to create, update, insert, or reference a Microsoft product icon deck organized by product family.

Typical requests:

- Create a Microsoft product icon deck.
- Add M365, D365, Azure, Power Platform, or Entra icons to a presentation.
- Generate a categorized icon catalog for Microsoft products.
- Reuse the public Microsoft product icon deck from GitHub.

## Primary Output

Use or generate this deck:

- Local public deck: `Microsoft_Product_Icons_By_Category_Public.pptx`
- GitHub repository: `https://github.com/AndyWujinhui/microsoft-product-icon-deck`
- GitHub deck path: `Microsoft_Product_Icons_By_Category_Public.pptx`

## Categories Included

- Microsoft 365
- Dynamics 365
- Azure - AI and Data
- Azure - App Platform
- Azure - Network, Security, Ops
- Power Platform
- Microsoft Entra

## Preferred Workflow

1. If the local project folder is available, use the existing deck first:

   ```powershell
   Microsoft_Product_Icons_By_Category_Public.pptx
   ```

2. If the deck needs to be regenerated, run:

   ```powershell
   npm install
   $env:PUBLIC_RELEASE='1'
   $env:OUTPUT_DECK='Microsoft_Product_Icons_By_Category_Public.pptx'
   npm run build
   Remove-Item Env:PUBLIC_RELEASE,Env:OUTPUT_DECK
   ```

3. Validate generated deck text:

   ```powershell
   python -m markitdown Microsoft_Product_Icons_By_Category_Public.pptx | Select-String -Pattern 'xxxx','lorem','ipsum','placeholder','Missing icons','not found','Generation note' -CaseSensitive:$false
   ```

4. If PowerPoint is available, export slide images for visual QA:

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

## Public Use Rule

For public or external decks, use only the public-release deck and set `PUBLIC_RELEASE=1` when regenerating. Do not include private roadmap decks, extracted roadmap media, `roadmap-deck-extracted/`, `roadmap-icon-candidates/`, `official-icon-assets/`, `generated-assets/`, `qa-slides/`, or `node_modules/` in any exported or published package.

## Icon Terms

Use Microsoft icons only under the Microsoft terms published on Microsoft Learn. Keep product names close to icons. Do not crop, flip, rotate, distort, or use Microsoft product icons to represent another product or service.
