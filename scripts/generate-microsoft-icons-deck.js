const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const sharp = require('sharp');
const ShapeType = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse' };

const root = path.resolve(__dirname, '..');
const assetsRoot = path.join(root, 'official-icon-assets', 'extracted');
const generatedRoot = path.join(root, 'generated-assets');
const outputDeck = path.join(root, process.env.OUTPUT_DECK || 'Microsoft_Product_Icons_By_Category.pptx');
const isPublicRelease = process.env.PUBLIC_RELEASE === '1';

const colors = {
  ink: '172033',
  muted: '5A6475',
  paper: 'F7F9FC',
  white: 'FFFFFF',
  azure: '0078D4',
  m365: '2F64C8',
  dynamics: '0E7878',
  power: '742774',
  entra: '4F6BED',
  fabric: 'E6A700',
  line: 'DCE3EE',
};

const sources = [
  ['Product roadmap deck', 'ERP product roadmap and innovation  2026-01.pptx'],
  ['Azure', 'https://learn.microsoft.com/en-us/azure/architecture/icons/'],
  ['Microsoft 365', 'https://learn.microsoft.com/en-us/microsoft-365/solutions/architecture-icons-templates'],
  ['Dynamics 365', 'https://learn.microsoft.com/en-us/dynamics365/get-started/icons'],
  ['Power Platform', 'https://learn.microsoft.com/en-us/power-platform/guidance/icons'],
  ['Microsoft Entra', 'https://learn.microsoft.com/en-us/entra/architecture/architecture-icons'],
].filter(([name]) => !isPublicRelease || name !== 'Product roadmap deck');

const categories = [
  {
    key: 'm365-official',
    title: 'Microsoft 365',
    subtitle: 'Official Microsoft 365 architecture symbols from Microsoft Learn',
    accent: colors.m365,
    root: 'Microsoft365',
    publicOnly: true,
    caption: 'Official Microsoft SVG',
    items: [
      { label: 'Teams', match: ['Teams Purple\\Building People_Teams_Dark.svg'] },
      { label: 'Teams meetings', match: ['Teams Purple\\48x48 Dark Purple Icon\\People Team.svg'] },
      { label: 'SharePoint sites', match: ['SharePoint Teal\\48x48 SVG Icon\\Organization_Dark.svg'] },
      { label: 'Planner', match: ['Planner Green\\48x48 SVG Icons\\Ribbon_Planner_Dark.svg'] },
      { label: 'Project', match: ['Project Green\\48x48 SVG Icons\\Document_Dark.svg'] },
      { label: 'Exchange mail', match: ['Microsoft Blue\\48x48 Dark Blue Icon\\Mail.svg'] },
      { label: 'Calendar', match: ['Microsoft Blue\\48x48 Dark Blue Icon\\Calendar Month.svg'] },
      { label: 'Compliance docs', match: ['Microsoft Blue\\48x48 Dark Blue Icon\\Document Lock.svg'] },
      { label: 'Cloud PC', match: ['Microsoft Blue\\48x48 Dark Blue Icon\\Cloud Desktop.svg'] },
      { label: 'Presentations', match: ['Microsoft Blue\\48x48 Dark Blue Icon\\Share Screen Person.svg'] },
    ],
  },
  {
    key: 'm365',
    title: 'Microsoft 365',
    subtitle: 'Product logos extracted from the added PG product roadmap deck',
    accent: colors.m365,
    root: 'Roadmap',
    privateOnly: true,
    caption: 'From PG roadmap deck',
    footer: 'Microsoft 365: product logos extracted from the added PG product roadmap deck.',
    items: [
      { label: 'Teams', match: ['Teams.png'] },
      { label: 'Word', match: ['Word.png'] },
      { label: 'PowerPoint', match: ['PowerPoint.png'] },
      { label: 'Excel', match: ['Excel.png'] },
      { label: 'Outlook', match: ['Outlook.png'] },
      { label: 'SharePoint', match: ['SharePoint.png'] },
      { label: 'Defender', match: ['Defender.png'] },
      { label: 'Windows 365', match: ['Windows365.png'] },
    ],
  },
  {
    key: 'roadmap-business-apps',
    title: 'Roadmap Deck Icons',
    subtitle: 'Additional product icons curated from the PG roadmap deck media',
    accent: colors.fabric,
    root: 'Roadmap',
    privateOnly: true,
    caption: 'From PG roadmap deck',
    footer: 'Roadmap Deck Icons: product logos extracted from the added PG product roadmap deck.',
    items: [
      { label: 'Power Apps', match: ['PowerApps.png'] },
      { label: 'Power Automate', match: ['PowerAutomate.png'] },
      { label: 'Dynamics 365', match: ['Dynamics365.png'] },
      { label: 'D365 Commerce', match: ['Dynamics365Commerce.png'] },
      { label: 'Teams', match: ['Teams.png'] },
      { label: 'Outlook', match: ['Outlook.png'] },
      { label: 'SharePoint', match: ['SharePoint.png'] },
      { label: 'Defender', match: ['Defender.png'] },
    ],
  },
  {
    key: 'dynamics',
    title: 'Dynamics 365',
    subtitle: 'Business applications for sales, service, finance, operations, commerce, and supply chain',
    accent: colors.dynamics,
    root: 'Dynamics365',
    caption: 'Official Microsoft SVG',
    items: [
      { label: 'Dynamics 365', match: ['Dynamics365_scalable.svg'] },
      { label: 'Sales', match: ['Sales_scalable.svg'] },
      { label: 'Customer Service', match: ['CustomerServices_scalable.svg'] },
      { label: 'Field Service', match: ['FieldService_scalable.svg'] },
      { label: 'Finance', match: ['Finance_scalable.svg'] },
      { label: 'Supply Chain', match: ['SupplyChainManagement_scalable.svg'] },
      { label: 'Business Central', match: ['BusinessCentral_scalable.svg'] },
      { label: 'Contact Center', match: ['ContactCenter_scalable.svg'] },
    ],
  },
  {
    key: 'azure-ai-data',
    title: 'Azure - AI and Data',
    subtitle: 'AI, machine learning, analytics, databases, and storage services',
    accent: colors.azure,
    root: 'Azure',
    caption: 'Official Microsoft SVG',
    items: [
      { label: 'Azure OpenAI', match: ['Azure-OpenAI.svg'] },
      { label: 'AI Studio', match: ['AI-Studio.svg'] },
      { label: 'Machine Learning', match: ['Machine-Learning.svg'] },
      { label: 'Data Factory', match: ['Data-Factories.svg'] },
      { label: 'Synapse Analytics', match: ['Azure-Synapse-Analytics.svg'] },
      { label: 'Databricks', match: ['Azure-Databricks.svg'] },
      { label: 'Azure SQL', match: ['Azure-SQL.svg', 'SQL-Database.svg'] },
      { label: 'Cosmos DB', match: ['Azure-Cosmos-DB.svg'] },
      { label: 'Storage Accounts', match: ['Storage-Accounts.svg'] },
    ],
  },
  {
    key: 'azure-app-platform',
    title: 'Azure - App Platform',
    subtitle: 'Compute, containers, web apps, API management, and integration services',
    accent: colors.azure,
    root: 'Azure',
    caption: 'Official Microsoft SVG',
    items: [
      { label: 'Virtual Machines', match: ['Virtual-Machines-(Classic).svg'] },
      { label: 'App Service', match: ['App-Services.svg'] },
      { label: 'Functions', match: ['Function-Apps.svg'] },
      { label: 'AKS', match: ['Kubernetes-Services.svg'] },
      { label: 'Container Apps', match: ['Container-Apps-Environments.svg'] },
      { label: 'Container Registry', match: ['Container-Registries.svg'] },
      { label: 'API Management', match: ['API-Management-Services.svg'] },
      { label: 'Logic Apps', match: ['Logic-Apps.svg'] },
      { label: 'Service Bus', match: ['Azure-Service-Bus.svg'] },
    ],
  },
  {
    key: 'azure-network-security',
    title: 'Azure - Network, Security, Ops',
    subtitle: 'Networking, secure access, threat protection, monitoring, and operations',
    accent: colors.azure,
    root: 'Azure',
    caption: 'Official Microsoft SVG',
    items: [
      { label: 'Virtual Network', match: ['Virtual-Networks.svg'] },
      { label: 'Application Gateway', match: ['Application-Gateways.svg'] },
      { label: 'Front Door', match: ['Front-Door-and-CDN-Profiles.svg'] },
      { label: 'Azure Firewall', match: ['Firewalls.svg'] },
      { label: 'Key Vault', match: ['Key-Vaults.svg'] },
      { label: 'Defender for Cloud', match: ['Microsoft-Defender-for-Cloud.svg'] },
      { label: 'Microsoft Sentinel', match: ['Azure-Sentinel.svg'] },
      { label: 'Event Hubs', match: ['Event-Hubs.svg'] },
      { label: 'Monitor', match: ['Monitor.svg'] },
      { label: 'Log Analytics', match: ['Log-Analytics-Workspaces.svg'] },
    ],
  },
  {
    key: 'power',
    title: 'Power Platform',
    subtitle: 'Low-code app development, automation, analytics, pages, Copilot Studio, and Dataverse',
    accent: colors.power,
    root: 'PowerPlatform',
    caption: 'Official Microsoft SVG',
    items: [
      { label: 'Power Platform', match: ['PowerPlatform_scalable.svg'] },
      { label: 'Power Apps', match: ['PowerApps_scalable.svg'] },
      { label: 'Power Automate', match: ['PowerAutomate_scalable.svg'] },
      { label: 'Power Pages', match: ['PowerPages_scalable.svg'] },
      { label: 'Copilot Studio', match: ['CopilotStudio_scalable.svg'] },
      { label: 'Dataverse', match: ['Dataverse_scalable.svg'] },
      { label: 'AI Builder', match: ['AIBuilder_scalable.svg'] },
      { label: 'Agent 365', match: ['Agent365_scalable.svg'] },
    ],
  },
  {
    key: 'entra',
    title: 'Microsoft Entra',
    subtitle: 'Identity, access, governance, verified identity, workload identity, and secure access',
    accent: colors.entra,
    root: 'Entra',
    caption: 'Official Microsoft SVG',
    items: [
      { label: 'Entra family', match: ['Microsoft Entra Product Family.svg'] },
      { label: 'Entra ID', match: ['Microsoft Entra ID color icon.svg'] },
      { label: 'ID Governance', match: ['Microsoft Entra ID Governance color icon.svg'] },
      { label: 'Internet Access', match: ['Microsoft Entra Internet Access color icon.svg'] },
      { label: 'Private Access', match: ['Microsoft Entra Private Access color icon.svg'] },
      { label: 'Verified ID', match: ['Microsoft Entra Verified ID color icon.svg'] },
      { label: 'Workload ID', match: ['Microsoft Entra Workload ID color icon.svg'] },
    ],
  },
].filter((category) => (isPublicRelease ? !category.privateOnly : !category.publicOnly));

function walkFiles(dir) {
  if (!fs.existsSync(dir)) return [];
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  return entries.flatMap((entry) => {
    const full = path.join(dir, entry.name);
    return entry.isDirectory() ? walkFiles(full) : [full];
  });
}

const allIconFiles = walkFiles(assetsRoot).filter((file) => /\.(svg|png)$/i.test(file));

function findIcon(category, item) {
  const categoryRoot = path.join(assetsRoot, category.root);
  const files = allIconFiles.filter((file) => file.startsWith(categoryRoot));
  for (const token of item.match) {
    const normalizedToken = token.toLowerCase().replace(/\\/g, path.sep).replace(/\//g, path.sep);
    const exactPath = files.find((file) => file.toLowerCase().endsWith(normalizedToken));
    if (exactPath) return exactPath;
    const exact = files.find((file) => path.basename(file).toLowerCase() === normalizedToken);
    if (exact) return exact;
    const partial = files.find((file) => path.basename(file).toLowerCase().includes(normalizedToken.replace(/\.svg|\.png/g, '').toLowerCase()));
    if (partial) return partial;
  }
  return null;
}

async function iconToPng(file, name) {
  if (!file) return null;
  fs.mkdirSync(generatedRoot, { recursive: true });
  const out = path.join(generatedRoot, `${name}.png`);
  if (/\.png$/i.test(file)) {
    fs.copyFileSync(file, out);
  } else {
    await sharp(file, { density: 240 }).resize(512, 512, { fit: 'inside', withoutEnlargement: false }).png().toFile(out);
  }
  return out;
}

function addFooter(slide, label) {
  slide.addText(label, {
    x: 0.55,
    y: 5.32,
    w: 8.9,
    h: 0.15,
    fontFace: 'Aptos',
    fontSize: 6.5,
    color: colors.muted,
    margin: 0,
  });
}

function addTopBar(slide, category) {
  slide.background = { color: colors.paper };
  slide.addShape(ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.16, fill: { color: category.accent }, line: { color: category.accent } });
  slide.addText(category.title, {
    x: 0.55,
    y: 0.38,
    w: 4.7,
    h: 0.43,
    fontFace: 'Aptos Display',
    fontSize: 26,
    bold: true,
    color: colors.ink,
    margin: 0,
  });
  slide.addText(category.subtitle, {
    x: 0.56,
    y: 0.86,
    w: 7.4,
    h: 0.32,
    fontFace: 'Aptos',
    fontSize: 9.5,
    color: colors.muted,
    margin: 0,
  });
}

function addIconCard(slide, item, iconPath, x, y, w, h, accent, caption) {
  const compact = h < 0.78;
  const iconSize = compact ? 0.5 : 0.58;
  const iconX = x + 0.18;
  const iconY = y + (compact ? 0.13 : 0.18);
  const textX = x + 0.9;
  slide.addShape(ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.04,
    fill: { color: colors.white },
    line: { color: colors.line, width: 0.8 },
    shadow: { type: 'outer', color: '8090A5', opacity: 0.13, blur: 1.5, angle: 45, offset: 1 },
  });
  slide.addShape(ShapeType.rect, { x, y, w: 0.06, h, fill: { color: accent }, line: { color: accent, transparency: 100 } });
  if (iconPath) {
    slide.addImage({ path: iconPath, x: iconX, y: iconY, w: iconSize, h: iconSize, altText: item.label });
  } else {
    slide.addShape(ShapeType.ellipse, { x: iconX + 0.04, y: iconY + 0.04, w: iconSize - 0.08, h: iconSize - 0.08, fill: { color: accent, transparency: 18 }, line: { color: accent, transparency: 100 } });
    slide.addText('?', { x: iconX + 0.04, y: iconY + 0.13, w: iconSize - 0.08, h: 0.18, align: 'center', fontSize: 13, bold: true, color: accent, margin: 0 });
  }
  slide.addText(item.label, {
    x: textX,
    y: y + (compact ? 0.18 : 0.22),
    w: w - 1.05,
    h: compact ? 0.32 : 0.43,
    fontFace: 'Aptos',
    fontSize: 11.5,
    bold: true,
    color: colors.ink,
    valign: 'mid',
    fit: 'shrink',
    margin: 0.02,
  });
  slide.addText(caption,
  {
    x: textX,
    y: y + h - 0.17,
    w: w - 1.05,
    h: 0.2,
    fontFace: 'Aptos',
    fontSize: 6.8,
    color: colors.muted,
    margin: 0,
  });
}

function addTitleSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: colors.ink };
  slide.addShape(ShapeType.rect, { x: 0, y: 0, w: 3.2, h: 5.625, fill: { color: '101827' }, line: { color: '101827' } });
  slide.addShape(ShapeType.rect, { x: 3.2, y: 0, w: 0.08, h: 5.625, fill: { color: colors.azure }, line: { color: colors.azure } });
  const tiles = [colors.azure, colors.dynamics, colors.power, colors.entra, colors.fabric];
  tiles.forEach((color, index) => {
    slide.addShape(ShapeType.roundRect, {
      x: 0.7 + (index % 2) * 0.92,
      y: 1.22 + Math.floor(index / 2) * 0.9,
      w: 0.64,
      h: 0.64,
      rectRadius: 0.04,
      fill: { color },
      line: { color, transparency: 100 },
    });
  });
  slide.addText('Microsoft Product Icons', {
    x: 3.85,
    y: 1.26,
    w: 5.2,
    h: 0.65,
    fontFace: 'Aptos Display',
    fontSize: 34,
    bold: true,
    color: colors.white,
    margin: 0,
  });
  slide.addText('按产品类别整理的官方图标 Deck', {
    x: 3.88,
    y: 2.02,
    w: 5.1,
    h: 0.35,
    fontFace: 'Microsoft YaHei UI',
    fontSize: 15,
    color: 'DDE7F7',
    margin: 0,
  });
  slide.addText('M365 / D365 / Azure / Power Platform / Entra', {
    x: 3.9,
    y: 2.72,
    w: 5.4,
    h: 0.32,
    fontFace: 'Aptos',
    fontSize: 12.5,
    color: 'AFC1D9',
    margin: 0,
  });
  slide.addText('Icons downloaded from Microsoft official Learn pages and official Microsoft-hosted repositories.', {
    x: 3.9,
    y: 4.72,
    w: 5.4,
    h: 0.28,
    fontFace: 'Aptos',
    fontSize: 8.5,
    color: 'AFC1D9',
    margin: 0,
  });
}

function addOverviewSlide(pres) {
  const slide = pres.addSlide();
  slide.background = { color: colors.paper };
  slide.addText('产品类别总览', {
    x: 0.55,
    y: 0.45,
    w: 4.8,
    h: 0.5,
    fontFace: 'Microsoft YaHei UI',
    fontSize: 26,
    bold: true,
    color: colors.ink,
    margin: 0,
  });
  slide.addText('每个类别页使用对应官方图标包中的 SVG，并保留产品名标签。', {
    x: 0.58,
    y: 0.98,
    w: 7.3,
    h: 0.28,
    fontFace: 'Microsoft YaHei UI',
    fontSize: 10,
    color: colors.muted,
    margin: 0,
  });
  categories.forEach((category, index) => {
    const x = 0.55 + (index % 4) * 2.35;
    const y = 1.72 + Math.floor(index / 4) * 1.25;
    slide.addShape(ShapeType.roundRect, { x, y, w: 2.02, h: 0.94, rectRadius: 0.04, fill: { color: colors.white }, line: { color: colors.line } });
    slide.addShape(ShapeType.rect, { x, y, w: 2.02, h: 0.12, fill: { color: category.accent }, line: { color: category.accent } });
    slide.addText(category.title, { x: x + 0.16, y: y + 0.25, w: 1.72, h: 0.34, fontFace: 'Aptos Display', fontSize: 12.4, bold: true, color: colors.ink, margin: 0, fit: 'shrink' });
    slide.addText(`${category.items.length} representative icons`, { x: x + 0.16, y: y + 0.66, w: 1.72, h: 0.18, fontFace: 'Aptos', fontSize: 7.2, color: colors.muted, margin: 0 });
  });
  addFooter(slide, 'Source: Microsoft Learn official icon pages. See appendix for URLs and permitted-use note.');
}

function addCategorySlide(pres, category, renderedItems) {
  const slide = pres.addSlide();
  addTopBar(slide, category);
  const cols = category.items.length > 8 ? 3 : 2;
  const rows = Math.ceil(category.items.length / cols);
  const cardW = cols === 3 ? 2.8 : 4.16;
  const cardH = rows >= 4 ? 0.72 : 0.82;
  const x0 = cols === 3 ? 0.55 : 0.68;
  const y0 = 1.48;
  const gapX = cols === 3 ? 0.25 : 0.5;
  const gapY = rows >= 4 ? 0.18 : 0.24;
  renderedItems.forEach((item, index) => {
    const col = index % cols;
    const row = Math.floor(index / cols);
    addIconCard(slide, item, item.renderedPath, x0 + col * (cardW + gapX), y0 + row * (cardH + gapY), cardW, cardH, category.accent, category.caption || 'Icon asset');
  });
  addFooter(slide, category.footer || `${category.title}: icons downloaded from Microsoft official icon package; do not crop, rotate, distort, or rebrand.`);
}

function addSourceSlide(pres, missing) {
  const slide = pres.addSlide();
  slide.background = { color: '101827' };
  slide.addText('官方来源与使用说明', {
    x: 0.55,
    y: 0.44,
    w: 5.5,
    h: 0.42,
    fontFace: 'Microsoft YaHei UI',
    fontSize: 23,
    bold: true,
    color: colors.white,
    margin: 0,
  });
  slide.addText('Microsoft permits these icons for architectural diagrams, training materials, and documentation. Keep product names close to icons; do not crop, flip, rotate, distort, or use the icons to represent your own product or service.', {
    x: 0.58,
    y: 1.03,
    w: 8.6,
    h: 0.52,
    fontFace: 'Aptos',
    fontSize: 10.5,
    color: 'DDE7F7',
    margin: 0,
    breakLine: false,
    fit: 'shrink',
  });
  sources.forEach(([name, url], index) => {
    const y = 1.92 + index * 0.48;
    slide.addText(name, { x: 0.72, y, w: 1.55, h: 0.18, fontFace: 'Aptos', fontSize: 9.5, bold: true, color: colors.white, margin: 0 });
    slide.addText(url, { x: 2.25, y, w: 6.9, h: 0.18, fontFace: 'Aptos', fontSize: 7.2, color: 'AFC1D9', margin: 0, fit: 'shrink' });
  });
  if (missing.length) {
    slide.addText(`Generation note: ${missing.length} requested icons were not found in the downloaded packages and were marked in-place.`, {
      x: 0.72,
      y: 4.72,
      w: 8.3,
      h: 0.25,
      fontFace: 'Aptos',
      fontSize: 8.2,
      color: 'FFD166',
      margin: 0,
    });
  }
}

async function main() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'GitHub Copilot';
  pptx.company = 'Microsoft';
  pptx.subject = 'Microsoft official product icons by category';
  pptx.title = 'Microsoft Product Icons By Category';
  pptx.lang = 'zh-CN';
  pptx.theme = {
    headFontFace: 'Aptos Display',
    bodyFontFace: 'Aptos',
    lang: 'zh-CN',
  };

  addTitleSlide(pptx);
  addOverviewSlide(pptx);

  const missing = [];
  for (const category of categories) {
    const renderedItems = [];
    for (const item of category.items) {
      const icon = findIcon(category, item);
      if (!icon) missing.push(`${category.title}: ${item.label}`);
      const renderedPath = await iconToPng(icon, `${category.key}-${item.label.replace(/[^a-z0-9]+/gi, '-').toLowerCase()}`);
      renderedItems.push({ ...item, renderedPath });
    }
    addCategorySlide(pptx, category, renderedItems);
  }

  addSourceSlide(pptx, missing);
  await pptx.writeFile({ fileName: outputDeck });

  console.log(`Wrote ${outputDeck}`);
  console.log(`Slides: ${pptx._slides.length}`);
  if (missing.length) {
    console.log('Missing icons:');
    missing.forEach((item) => console.log(`- ${item}`));
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});