import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const root = new URL("../", import.meta.url);

test("the landing page links its local assets and main sections", async () => {
  const html = await readFile(new URL("index.html", root), "utf8");

  assert.match(html, /<html lang="ru">/);
  assert.match(html, /href="styles\.css"/);
  assert.match(html, /src="script\.js"/);
  assert.match(html, /<main id="top">/);
  assert.match(html, /id="about"/);
  assert.match(html, /id="contact"/);
});

test("all referenced local assets exist and contain content", async () => {
  for (const asset of ["styles.css", "script.js"]) {
    const contents = await readFile(new URL(asset, root), "utf8");
    assert.ok(contents.trim().length > 0, `${asset} should not be empty`);
  }
});
