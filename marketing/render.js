// Renders the store/social images: node marketing/render.js
const { chromium } = require("playwright");
const path = require("path");
const page = "file://" + path.resolve(__dirname, "cover.html");
const shots = [
  ["gumroad-cover-1280x720.png", 1280, 720],
  ["gumroad-thumb-600x600.png", 600, 600],
  ["etsy-listing-2000x1500.png", 2000, 1500],
  ["pinterest-pin-1000x1500.png", 1000, 1500],
  ["pinterest-pin-2-1000x1500.png", 1000, 1500, "An extra $100 a month saved $1,743 in interest"],
];
(async () => {
  const b = await chromium.launch();
  for (const [name, w, h, t] of shots) {
    const p = await b.newPage({ viewport: { width: w, height: h } });
    await p.goto(`${page}?w=${w}&h=${h}` + (t ? `&t=${encodeURIComponent(t)}` : ""));
    await p.screenshot({ path: path.join(__dirname, "images", name) });
    await p.close();
  }
  await b.close();
})();
