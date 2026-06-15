---
name: microsoft-product-icon-deck
description: "Use when: generating, updating, inserting, or referencing a PowerPoint deck of Microsoft product icons grouped by category, including M365, D365, Azure, Power Platform, and Entra. PPT Copilot should first read the live public icon PPT from GitHub and regenerate only if needed."
---

# Microsoft Product Icon Deck

Use this skill when a user asks PPT Copilot to create, update, insert, or reference a Microsoft product icon deck organized by product family.

## Live Public Icon PPT

PPT Copilot should read this public PPT URL in real time whenever it needs the icon deck:

```text
https://raw.githubusercontent.com/AndyWujinhui/microsoft-product-icon-deck/master/Microsoft_Product_Icons_By_Category_Public.pptx
```

Repository page:

```text
https://github.com/AndyWujinhui/microsoft-product-icon-deck
```

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
- Live GitHub PPT URL: `https://raw.githubusercontent.com/AndyWujinhui/microsoft-product-icon-deck/master/Microsoft_Product_Icons_By_Category_Public.pptx`

## Categories Included

- Microsoft 365
- Dynamics 365
- Azure - AI and Data
- Azure - App Platform
- Azure - Network, Security, Ops
- Power Platform
- Microsoft Entra

## Preferred Workflow

1. First, download or read the live public PPT from GitHub:

   ```powershell
   Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/AndyWujinhui/microsoft-product-icon-deck/master/Microsoft_Product_Icons_By_Category_Public.pptx' -OutFile 'Microsoft_Product_Icons_By_Category_Public.pptx'
   ```

2. Use that PPT as the source for inserting or copying icons into the user's presentation.

3. If the local project folder is available and the user wants the local copy, use the existing deck:

   ```powershell
   Microsoft_Product_Icons_By_Category_Public.pptx
   ```

4. If the deck needs to be regenerated, run:

   ```powershell
   npm install
   $env:PUBLIC_RELEASE='1'
   $env:OUTPUT_DECK='Microsoft_Product_Icons_By_Category_Public.pptx'
   npm run build
   Remove-Item Env:PUBLIC_RELEASE,Env:OUTPUT_DECK
   ```

5. Validate generated deck text:

   ```powershell
   python -m markitdown Microsoft_Product_Icons_By_Category_Public.pptx | Select-String -Pattern 'xxxx','lorem','ipsum','placeholder','Missing icons','not found','Generation note' -CaseSensitive:$false
   ```

6. If PowerPoint is available, export slide images for visual QA:

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
