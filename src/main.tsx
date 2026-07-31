import React from 'react';
import { createRoot } from 'react-dom/client';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';
import JSZip from 'jszip';
import './styles.css';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

type Asset = {
  id: string;
  number: number;
  name: string;
  source: string;
  method: string;
  blob: Blob;
  url: string;
  thumb: string;
  width: number;
  height: number;
  size: number;
  format: string;
  filtered: boolean;
  reason?: string;
  hash: string;
  edited: { name: string; blob: Blob }[];
};

const allowed = new Set(['pdf', 'docx', 'png', 'jpg', 'jpeg', 'webp', 'tif', 'tiff']);

function ext(name: string) {
  return (name.split('.').pop() || '').toLowerCase();
}

function formatBytes(n: number) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${Math.round(n / 1024)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

async function sha256(blob: Blob) {
  const digest = await crypto.subtle.digest('SHA-256', await blob.arrayBuffer());
  return Array.from(new Uint8Array(digest))
    .map((b) => b.toString(16).padStart(2, '0'))
    .join('');
}

function canvasToBlob(canvas: HTMLCanvasElement, type = 'image/png', quality = 0.92): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error('Could not export image'));
    }, type, quality);
  });
}

function blobToImage(blob: Blob): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const url = URL.createObjectURL(blob);

    img.onload = () => {
      URL.revokeObjectURL(url);
      resolve(img);
    };

    img.onerror = reject;
    img.src = url;
  });
}

async function getDimensions(blob: Blob) {
  const img = await blobToImage(blob);
  return {
    width: img.naturalWidth,
    height: img.naturalHeight
  };
}

async function makeThumbnail(blob: Blob) {
  const img = await blobToImage(blob);
  const scale = Math.min(360 / img.naturalWidth, 260 / img.naturalHeight, 1);

  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(img.naturalWidth * scale));
  canvas.height = Math.max(1, Math.round(img.naturalHeight * scale));

  canvas.getContext('2d')!.drawImage(img, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL('image/png');
}

async function isBlank(blob: Blob) {
  const img = await blobToImage(blob);

  const canvas = document.createElement('canvas');
  const scale = Math.min(120 / img.naturalWidth, 120 / img.naturalHeight, 1);

  canvas.width = Math.max(1, Math.round(img.naturalWidth * scale));
  canvas.height = Math.max(1, Math.round(img.naturalHeight * scale));

  const ctx = canvas.getContext('2d', { willReadFrequently: true })!;
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

  const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;

  let sum = 0;
  let sum2 = 0;
  let count = 0;

  for (let i = 0; i < data.length; i += 4) {
    const grey = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
    sum += grey;
    sum2 += grey * grey;
    count++;
  }

  const variance = sum2 / count - Math.pow(sum / count, 2);
  return variance < 2;
}

async function makeAsset(
  blob: Blob,
  source: string,
  method: string,
  format: string,
  current: Asset[]
): Promise<Asset> {
  const dimensions = await getDimensions(blob);
  const hash = await sha256(blob);
  const duplicate = current.find((asset) => asset.hash === hash);
  const blank = await isBlank(blob);
  const tiny = dimensions.width * dimensions.height < 10000;
  const number = current.length + 1;

  const filtered = blank || tiny || Boolean(duplicate);
  const reason = blank
    ? 'Blank or low visual signal'
    : tiny
      ? 'Tiny decorative asset'
      : duplicate
        ? `Duplicate of Image ${duplicate.number}`
        : undefined;

  return {
    id: crypto.randomUUID(),
    number,
    name: `image-${String(number).padStart(3, '0')}.${format.toLowerCase()}`,
    source,
    method,
    blob,
    url: URL.createObjectURL(blob),
    thumb: await makeThumbnail(blob),
    width: dimensions.width,
    height: dimensions.height,
    size: blob.size,
    format: format.toUpperCase(),
    filtered,
    reason,
    hash,
    edited: []
  };
}

async function extractDocx(file: File) {
  const zip = await JSZip.loadAsync(file);
  const results: { blob: Blob; source: string; method: string; format: string }[] = [];

  for (const item of Object.values(zip.files)) {
    if (item.dir) continue;
    if (!/^word\/media\//.test(item.name)) continue;

    const fileExt = ext(item.name);

    if (!['png', 'jpg', 'jpeg', 'webp', 'tif', 'tiff', 'bmp', 'gif'].includes(fileExt)) {
      continue;
    }

    results.push({
      blob: await item.async('blob'),
      source: `DOCX ${item.name}`,
      method: 'docx-original-image',
      format: fileExt
    });
  }

  return results;
}

async function extractPdf(file: File) {
  const pdf = await pdfjsLib.getDocument({
    data: await file.arrayBuffer()
  }).promise;

  const results: { blob: Blob; source: string; method: string; format: string }[] = [];

  for (let pageNo = 1; pageNo <= pdf.numPages; pageNo++) {
    const page = await pdf.getPage(pageNo);
    const viewport = page.getViewport({ scale: 2.5 });

    const canvas = document.createElement('canvas');
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    await page.render({
      canvasContext: canvas.getContext('2d')!,
      viewport
    }).promise;

    results.push({
      blob: await canvasToBlob(canvas),
      source: `PDF page ${pageNo}`,
      method: 'pdf-rendered-page',
      format: 'png'
    });
  }

  return results;
}

function download(blob: Blob, name: string) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();

  setTimeout(() => {
    URL.revokeObjectURL(a.href);
  }, 5000);
}

function Editor({
  asset,
  total,
  onClose,
  onSave
}: {
  asset: Asset;
  total: number;
  onClose: () => void;
  onSave: (blob: Blob, extension: string) => void;
}) {
  const canvasRef = React.useRef<HTMLCanvasElement>(null);
  const [format, setFormat] = React.useState('image/png');
  const [text, setText] = React.useState('Note');

  React.useEffect(() => {
    async function load() {
      const img = await blobToImage(asset.blob);
      const canvas = canvasRef.current!;

      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;

      canvas.getContext('2d')!.drawImage(img, 0, 0);
    }

    load();
  }, [asset.id]);

  function rotateRight() {
    const canvas = canvasRef.current!;
    const source = document.createElement('canvas');

    source.width = canvas.width;
    source.height = canvas.height;
    source.getContext('2d')!.drawImage(canvas, 0, 0);

    canvas.width = source.height;
    canvas.height = source.width;

    const ctx = canvas.getContext('2d')!;
    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate(Math.PI / 2);
    ctx.drawImage(source, -source.width / 2, -source.height / 2);
  }

  function cropCentre() {
    const canvas = canvasRef.current!;
    const width = Math.round(canvas.width * 0.75);
    const height = Math.round(canvas.height * 0.75);
    const x = Math.round((canvas.width - width) / 2);
    const y = Math.round((canvas.height - height) / 2);

    const crop = document.createElement('canvas');
    crop.width = width;
    crop.height = height;

    crop.getContext('2d')!.drawImage(canvas, x, y, width, height, 0, 0, width, height);

    canvas.width = width;
    canvas.height = height;
    canvas.getContext('2d')!.drawImage(crop, 0, 0);
  }

  function applyFilter(filter: string) {
    const canvas = canvasRef.current!;
    const copy = document.createElement('canvas');

    copy.width = canvas.width;
    copy.height = canvas.height;

    const copyCtx = copy.getContext('2d')!;
    copyCtx.filter = filter;
    copyCtx.drawImage(canvas, 0, 0);

    const ctx = canvas.getContext('2d')!;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(copy, 0, 0);
  }

  function addText() {
    const canvas = canvasRef.current!;
    const ctx = canvas.getContext('2d')!;

    ctx.font = `${Math.max(24, canvas.width * 0.04)}px system-ui`;
    ctx.fillStyle = '#1457c8';
    ctx.fillText(text, 40, 60);
  }

  async function exportImage(saveVersion: boolean) {
    const extension = format === 'image/png' ? 'png' : format === 'image/jpeg' ? 'jpg' : 'webp';
    const blob = await canvasToBlob(canvasRef.current!, format, 0.92);

    if (saveVersion) {
      onSave(blob, extension);
    } else {
      download(blob, `image-${String(asset.number).padStart(3, '0')}-edited.${extension}`);
    }
  }

  return (
    <div className="backdrop" role="dialog" aria-modal="true">
      <section className="editor">
        <header>
          <h2>Editing Image {asset.number} of {total}</h2>
          <button onClick={onClose}>Close</button>
        </header>

        <div className="editGrid">
          <div className="canvasBox">
            <canvas ref={canvasRef} />
          </div>

          <aside>
            <button onClick={rotateRight}>Rotate right</button>
            <button onClick={cropCentre}>Crop centre</button>
            <button onClick={() => applyFilter('brightness(1.2)')}>Brighten</button>
            <button onClick={() => applyFilter('contrast(1.25)')}>Increase contrast</button>
            <button onClick={() => applyFilter('grayscale(1)')}>Greyscale</button>

            <label>
              Text
              <input value={text} onChange={(event) => setText(event.target.value)} />
            </label>

            <button onClick={addText}>Add text</button>

            <label>
              Export format
              <select value={format} onChange={(event) => setFormat(event.target.value)}>
                <option value="image/png">PNG</option>
                <option value="image/jpeg">JPG</option>
                <option value="image/webp">WEBP</option>
              </select>
            </label>

            <button onClick={() => exportImage(true)}>Save edited version</button>
            <button onClick={() => exportImage(false)}>Download edited</button>
          </aside>
        </div>
      </section>
    </div>
  );
}

function App() {
  const [assets, setAssets] = React.useState<Asset[]>([]);
  const [status, setStatus] = React.useState('Ready. Upload a PDF, DOCX or image file.');
  const [editing, setEditing] = React.useState<Asset | null>(null);
  const [selected, setSelected] = React.useState<Set<string>>(new Set());

  const activeAssets = assets.filter((asset) => !asset.filtered);

  async function processFile(file: File) {
    const fileExt = ext(file.name);

    if (fileExt === 'doc') {
      setStatus('Legacy .doc files are not supported. Save the file as DOCX or PDF first.');
      return;
    }

    if (!allowed.has(fileExt)) {
      setStatus('Unsupported file type. Upload PDF, DOCX, PNG, JPG, JPEG, WEBP or TIFF.');
      return;
    }

    setStatus('Extracting images...');

    try {
      const raw =
        fileExt === 'pdf'
          ? await extractPdf(file)
          : fileExt === 'docx'
            ? await extractDocx(file)
            : [
                {
                  blob: file,
                  source: 'Uploaded image',
                  method: 'uploaded-image',
                  format: fileExt
                }
              ];

      const next: Asset[] = [];

      for (const item of raw) {
        next.push(await makeAsset(item.blob, item.source, item.method, item.format, next));
      }

      setAssets(next);
      setSelected(new Set());

      setStatus(
        `Done. Extracted ${next.length} asset(s), ${next.filter((asset) => !asset.filtered).length} active.`
      );
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'Extraction failed.');
    }
  }

  async function downloadZip(items: Asset[]) {
    const zip = new JSZip();

    for (const asset of items) {
      zip.file(asset.name, asset.blob);

      for (const version of asset.edited) {
        zip.file(version.name, version.blob);
      }
    }

    const blob = await zip.generateAsync({ type: 'blob' });
    download(blob, 'document-image-studio-images.zip');
  }

  function toggleSelected(id: string, checked: boolean) {
    setSelected((current) => {
      const next = new Set(current);

      if (checked) next.add(id);
      else next.delete(id);

      return next;
    });
  }

  function saveVersion(asset: Asset, blob: Blob, extension: string) {
    setAssets((list) =>
      list.map((item) =>
        item.id === asset.id
          ? {
              ...item,
              edited: [
                ...item.edited,
                {
                  name: `image-${String(item.number).padStart(3, '0')}-edited-v${item.edited.length + 1}.${extension}`,
                  blob
                }
              ]
            }
          : item
      )
    );

    setStatus(`Saved edited version for Image ${asset.number}.`);
  }

  return (
    <main>
      <header className="hero">
        <p className="eyebrow">React + Vercel</p>
        <h1>Document Image Studio</h1>
        <p>
          Upload a PDF, DOCX or image, extract visuals, select by number, edit and download.
          Files stay in your browser.
        </p>
      </header>

      <section className="panel upload">
        <input
          type="file"
          accept=".pdf,.docx,.png,.jpg,.jpeg,.webp,.tif,.tiff"
          onChange={(event) => {
            const file = event.target.files?.[0];
            if (file) processFile(file);
          }}
        />

        <p role="status">{status}</p>
      </section>

      {assets.length > 0 && (
        <section className="panel">
          <div className="toolbar">
            <button onClick={() => setSelected(new Set(activeAssets.map((asset) => asset.id)))}>
              Select all active
            </button>

            <button onClick={() => downloadZip(assets.filter((asset) => selected.has(asset.id)))}>
              Download selected ZIP
            </button>

            <button onClick={() => downloadZip(activeAssets)}>Download all active ZIP</button>
          </div>

          <div className="gallery">
            {assets.map((asset) => (
              <article key={asset.id} className={asset.filtered ? 'card filtered' : 'card'}>
                <strong>Image {asset.number}</strong>

                <label>
                  <input
                    type="checkbox"
                    checked={selected.has(asset.id)}
                    onChange={(event) => toggleSelected(asset.id, event.target.checked)}
                  />
                  Select
                </label>

                <button className="thumb" onClick={() => setEditing(asset)}>
                  <img
                    src={asset.thumb}
                    alt={`Extracted image ${asset.number}, ${asset.source}, ${asset.width} by ${asset.height} pixels`}
                  />
                </button>

                <dl>
                  <dt>Source</dt>
                  <dd>{asset.source}</dd>

                  <dt>Size</dt>
                  <dd>
                    {asset.width} x {asset.height}, {formatBytes(asset.size)}
                  </dd>

                  <dt>Method</dt>
                  <dd>{asset.method}</dd>

                  <dt>Format</dt>
                  <dd>{asset.format}</dd>

                  {asset.filtered && (
                    <>
                      <dt>Filtered</dt>
                      <dd>
                        {asset.reason}{' '}
                        <button
                          onClick={() =>
                            setAssets((list) =>
                              list.map((item) =>
                                item.id === asset.id
                                  ? { ...item, filtered: false, reason: undefined }
                                  : item
                              )
                            )
                          }
                        >
                          Restore
                        </button>
                      </dd>
                    </>
                  )}
                </dl>

                <button onClick={() => setEditing(asset)}>Edit</button>
                <button onClick={() => download(asset.blob, asset.name)}>Download original</button>
              </article>
            ))}
          </div>
        </section>
      )}

      {editing && (
        <Editor
          asset={editing}
          total={assets.length}
          onClose={() => setEditing(null)}
          onSave={(blob, extension) => saveVersion(editing, blob, extension)}
        />
      )}
    </main>
  );
}

createRoot(document.getElementById('root')!).render(<App />);
