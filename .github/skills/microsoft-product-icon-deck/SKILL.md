---
name: microsoft-product-icon-deck
description: "Use when: creating, updating, inserting, or referencing a PowerPoint deck of Microsoft product icons by category, including M365, D365, Azure, Power Platform, and Entra icons. PPT Copilot should first read the live public icon PPT from GitHub, then regenerate only if needed."
---

# Microsoft Product Icon Deck Skill

Use this skill when the user asks for a PowerPoint icon deck, Microsoft product icon catalog, or a categorized deck for Microsoft product families such as M365, D365, Azure, Power Platform, or Entra.

## Live Public Icon Deck

PPT Copilot should read this live public PPT first whenever it needs Microsoft product icons:

```text
https://raw.githubusercontent.com/AndyWujinhui/microsoft-product-icon-deck/master/Microsoft_Product_Icons_By_Category_Public.pptx
```

Repository page:

```text
https://github.com/AndyWujinhui/microsoft-product-icon-deck
```

## Output

Use, download, or update `Microsoft_Product_Icons_By_Category_Public.pptx`.

## Workflow

1. Prefer reading the live public deck directly from GitHub:

   ```powershell
   Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/AndyWujinhui/microsoft-product-icon-deck/master/Microsoft_Product_Icons_By_Category_Public.pptx' -OutFile 'Microsoft_Product_Icons_By_Category_Public.pptx'
   ```

2. Use that PPT as the source deck for inserting or copying icons into the user's presentation.
3. Regenerate only when the user asks to update the catalog itself. Use the repository script `scripts/generate-microsoft-icons-deck.js`.
4. Build the public deck with:

   ```powershell
   $env:PUBLIC_RELEASE='1'
   $env:OUTPUT_DECK='Microsoft_Product_Icons_By_Category_Public.pptx'
   npm run build
   Remove-Item Env:PUBLIC_RELEASE,Env:OUTPUT_DECK
   ```

5. Validate the generated deck content:

   ```powershell
   python -m markitdown Microsoft_Product_Icons_By_Category_Public.pptx | Select-String -Pattern 'xxxx','lorem','ipsum','placeholder','Missing icons','not found','Generation note' -CaseSensitive:$false
   ```

6. If PowerPoint is available, export slides to PNG for visual QA:

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
