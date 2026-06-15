# Microsoft Product Icons By Category Deck

This repository contains a generated PowerPoint deck that organizes Microsoft product and architecture icons by product family.

## Deck

- `Microsoft_Product_Icons_By_Category_Public.pptx`

The public deck is generated only from Microsoft public icon sources, including Microsoft Learn icon packages for Azure, Microsoft 365, Dynamics 365, Power Platform, and Microsoft Entra.

## Categories

- Microsoft 365
- Dynamics 365
- Azure - AI and Data
- Azure - App Platform
- Azure - Network, Security, Ops
- Power Platform
- Microsoft Entra

## Generate

Install dependencies:

```powershell
npm install
```

Build the public deck:

```powershell
$env:PUBLIC_RELEASE='1'
$env:OUTPUT_DECK='Microsoft_Product_Icons_By_Category_Public.pptx'
npm run build
Remove-Item Env:PUBLIC_RELEASE,Env:OUTPUT_DECK
```

The script expects the public Microsoft icon packages to be available under `official-icon-assets/extracted`. The generated public deck is checked in so viewers do not need to run the build.

## Icon Terms

Microsoft permits these icons for architectural diagrams, training materials, and documentation under the terms published on the corresponding Microsoft Learn icon pages. Follow the Microsoft guidance: keep product names near icons, do not crop, flip, rotate, distort, or use Microsoft product icons to represent your own product or service.

## Sources

- Azure icons: https://learn.microsoft.com/en-us/azure/architecture/icons/
- Microsoft 365 architecture icons: https://learn.microsoft.com/en-us/microsoft-365/solutions/architecture-icons-templates
- Dynamics 365 icons: https://learn.microsoft.com/en-us/dynamics365/get-started/icons
- Power Platform icons: https://learn.microsoft.com/en-us/power-platform/guidance/icons
- Microsoft Entra icons: https://learn.microsoft.com/en-us/entra/architecture/architecture-icons
