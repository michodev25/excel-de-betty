const fs = require('node:fs/promises');
const sharp = require('C:/Users/Micho/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/sharp');

async function main() {
  const root = 'C:/Users/Micho/Desktop/programando/excel de betty';
  const original = 'C:/Users/Micho/AppData/Local/Temp/codex-clipboard-1307fcf6-a3c4-4d25-8ce4-143c70264d79.png';
  const qr = 'C:/Users/Micho/AppData/Local/Temp/codex-clipboard-00b43acc-99cc-4b0e-877e-ab3027a15eb5.png';
  const designed = 'C:/Users/Micho/.codex/generated_images/01a0e505-e0f3-7810-b846-ba709b231109/exec-679023a5-3415-4d03-8211-fb6e2fa00edf.png';
  const outputDir = root + '/output/yeya';
  await fs.mkdir(outputDir, { recursive: true });
  const topHeight = 1086;
  const footerHeight = 316;
  const top = await sharp(original).extract({ left: 0, top: 0, width: 853, height: topHeight }).toBuffer();
  const footer = await sharp(designed).extract({ left: 0, top: 1207, width: 1024, height: 329 }).resize(853, footerHeight, { fit: 'fill' }).toBuffer();
  const exactQr = await sharp(qr).resize(250, 250, { kernel: 'nearest' }).toBuffer();
  const output = outputDir + '/flyer-yeya-con-qr.png';
  await sharp({ create: { width: 853, height: topHeight + footerHeight, channels: 3, background: '#071a36' } })
    .composite([
      { input: top, left: 0, top: 0 },
      { input: footer, left: 0, top: topHeight },
      { input: exactQr, left: 563, top: topHeight + 49 },
    ])
    .png()
    .toFile(output);
  const sourcePixels = await sharp(original).extract({ left: 0, top: 0, width: 853, height: topHeight }).removeAlpha().raw().toBuffer();
  const outputPixels = await sharp(output).extract({ left: 0, top: 0, width: 853, height: topHeight }).removeAlpha().raw().toBuffer();
  const qrPixels = await sharp(exactQr).removeAlpha().raw().toBuffer();
  const outputQrPixels = await sharp(output).extract({ left: 563, top: topHeight + 49, width: 250, height: 250 }).removeAlpha().raw().toBuffer();
  const report = { output, dimensions: [853, topHeight + footerHeight], originalLogoAndInformationPixelIdentical: sourcePixels.equals(outputPixels), insertedQrMatchesSourceResizedWithNearestNeighbor: qrPixels.equals(outputQrPixels), qrBox: { x: 563, y: topHeight + 49, width: 250, height: 250 } };
  await fs.writeFile(outputDir + '/verificacion.json', JSON.stringify(report, null, 2));
  console.log(JSON.stringify(report));
}
main().catch(error => { console.error(error); process.exitCode = 1; });
