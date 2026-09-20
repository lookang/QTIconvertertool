async (page) => {
  const results = await page.evaluate(async () => {
    const root = new URL('files/doc/', location.href);

    async function listDocxFiles(directoryUrl) {
      const response = await fetch(directoryUrl);
      if (!response.ok) throw new Error(`Cannot list ${directoryUrl}: HTTP ${response.status}`);
      const html = await response.text();
      const listing = new DOMParser().parseFromString(html, 'text/html');
      const files = [];
      for (const anchor of listing.querySelectorAll('a[href]')) {
        const child = new URL(anchor.getAttribute('href'), directoryUrl);
        if (!child.href.startsWith(root.href) || child.href === directoryUrl) continue;
        if (child.pathname.endsWith('/')) {
          files.push(...await listDocxFiles(child.href));
        } else if (/\.docx$/i.test(child.pathname)) {
          files.push(child.href);
        }
      }
      return files;
    }

    const docxUrls = [...new Set(await listDocxFiles(root.href))].sort();
    const rows = [];
    for (const url of docxUrls) {
      const name = decodeURIComponent(new URL(url).pathname.split('/').pop());
      const isMarkScheme = /(?:\bMS\b|mark.?scheme|answers?(?: only| key)?)/i.test(name);
      try {
        const response = await fetch(url);
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const file = new File([await response.blob()], name, {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
        if (isMarkScheme) {
          const answers = await parseMarkSchemeEnhanced(file);
          rows.push({ name, kind: 'mark-scheme', answers: Object.keys(answers).length, ok: true });
        } else {
          const parsed = await parseDocxToQuestions(file, {});
          const qtiBlob = await buildQtiZipBlob(parsed.allQuestions, parsed.resolvedImages);
          const qtiZip = await JSZip.loadAsync(qtiBlob);
          const qtiNames = Object.keys(qtiZip.files);
          const itemNames = qtiNames.filter(path => /^items\/[^/]+\.xml$/i.test(path));
          const assetNames = qtiNames.filter(path => path.startsWith('assets/') && !qtiZip.files[path].dir);
          const itemXml = await Promise.all(itemNames.map(path => qtiZip.file(path).async('string')));
          const imageReferences = itemXml.flatMap(xml =>
            Array.from(xml.matchAll(/src="\.\.\/(assets\/[^"?#]+)"/g), match => match[1])
          );
          const missingImageAssets = imageReferences.filter(path => !assetNames.includes(path));
          const expectedImages = Object.keys(parsed.resolvedImages).length;
          const packageValid = itemNames.length === parsed.allQuestions.length
            && assetNames.length === expectedImages
            && missingImageAssets.length === 0
            && !!qtiZip.file('imsmanifest.xml')
            && !!qtiZip.file('assessment_test.xml');
          rows.push({
            name,
            kind: 'question-paper',
            questions: parsed.allQuestions.length,
            images: expectedImages,
            qtiItems: itemNames.length,
            qtiAssets: assetNames.length,
            packageValid,
            ok: parsed.allQuestions.length > 0 && packageValid
          });
        }
      } catch (error) {
        rows.push({ name, kind: isMarkScheme ? 'mark-scheme' : 'question-paper', ok: false, error: error.message });
      }
    }
    return rows;
  });

  await page.evaluate(rows => { window.__qtiBatchRegressionResults = rows; }, results);
}
