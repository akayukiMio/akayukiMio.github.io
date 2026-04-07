'use strict';

const crypto = require('crypto');
const fs = require('fs');
const fsp = fs.promises;
const path = require('path');

const chokidar = require('chokidar');
const mammoth = require('mammoth');
const TurndownService = require('turndown');

const SOURCE_FOLDER_NAME = '大澪神神狂想集！';
const MANAGED_POST_DIR = path.join('source', '_posts', '大澪神神狂想集');
const MANAGED_ASSET_DIR = path.join('source', 'docx-assets', '大澪神神狂想集');
const MANIFEST_PATH = path.join('.cache', 'docx-essay-manifest.json');
const TAG_NAME = '大澪神神狂想集';
const PUBLIC_ASSET_PREFIX = '/docx-assets/大澪神神狂想集';

const turndownService = new TurndownService({
  headingStyle: 'atx',
  bulletListMarker: '-',
  codeBlockStyle: 'fenced',
  emDelimiter: '*'
});

turndownService.remove(['style', 'script']);

function normalizePath(filePath) {
  return filePath.split(path.sep).join('/');
}

function pad(value) {
  return String(value).padStart(2, '0');
}

function formatDate(value) {
  const date = value instanceof Date ? value : new Date(value);
  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate())
  ].join('-') + ' ' + [
    pad(date.getHours()),
    pad(date.getMinutes()),
    pad(date.getSeconds())
  ].join(':');
}

function escapeYamlString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function buildEntryId(relativePath) {
  return `essay-${crypto.createHash('sha1').update(relativePath).digest('hex').slice(0, 16)}`;
}

function sanitizeSlugPart(value) {
  return String(value)
    .normalize('NFKC')
    .replace(/[\\/:*?"<>|#%&{}]+/g, ' ')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 80) || 'article';
}

function buildPostFileName(title, entryId) {
  const shortHash = entryId.replace(/^essay-/, '').slice(0, 8);
  return `${sanitizeSlugPart(title)}-${shortHash}.md`;
}

function getExtensionFromContentType(contentType) {
  const map = {
    'image/png': 'png',
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/gif': 'gif',
    'image/webp': 'webp',
    'image/svg+xml': 'svg',
    'image/bmp': 'bmp',
    'image/tiff': 'tif'
  };

  return map[contentType] || 'bin';
}

async function exists(targetPath) {
  try {
    await fsp.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

async function ensureDir(targetPath) {
  await fsp.mkdir(targetPath, { recursive: true });
}

async function removePath(targetPath) {
  await fsp.rm(targetPath, { recursive: true, force: true });
}

async function readJson(filePath, fallbackValue) {
  try {
    const content = await fsp.readFile(filePath, 'utf8');
    return JSON.parse(content);
  } catch (error) {
    if (error.code === 'ENOENT') {
      return fallbackValue;
    }

    throw error;
  }
}

async function writeFileIfChanged(filePath, content) {
  let previousContent = null;

  try {
    previousContent = await fsp.readFile(filePath, 'utf8');
  } catch (error) {
    if (error.code !== 'ENOENT') {
      throw error;
    }
  }

  if (previousContent === content) {
    return false;
  }

  await ensureDir(path.dirname(filePath));
  await fsp.writeFile(filePath, content, 'utf8');
  return true;
}

async function writeJsonIfChanged(filePath, data) {
  const content = `${JSON.stringify(data, null, 2)}\n`;
  return writeFileIfChanged(filePath, content);
}

async function listDocxFiles(rootDir) {
  if (!(await exists(rootDir))) {
    return [];
  }

  const results = [];

  async function walk(currentDir) {
    const entries = await fsp.readdir(currentDir, { withFileTypes: true });

    for (const entry of entries) {
      const fullPath = path.join(currentDir, entry.name);

      if (entry.isDirectory()) {
        await walk(fullPath);
        continue;
      }

      if (entry.isFile() && /\.docx$/i.test(entry.name)) {
        results.push(fullPath);
      }
    }
  }

  await walk(rootDir);
  results.sort((left, right) => left.localeCompare(right, 'zh-CN'));
  return results;
}

function normalizeMarkdown(markdown) {
  return markdown
    .replace(/\r\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function buildMarkdownDocument({ title, date, updated, body }) {
  const normalizedBody = body || '';

  return [
    '---',
    `title: ${escapeYamlString(title)}`,
    `date: ${date}`,
    `updated: ${updated}`,
    'tags:',
    `  - ${TAG_NAME}`,
    '---',
    '',
    normalizedBody,
    normalizedBody ? '' : ''
  ].join('\n');
}

async function convertDocxToManagedMarkdown({ sourceFilePath, relativePath, entryId, title, firstImportedAt, updatedAt, assetRootDir, tempAssetRootDir }) {
  const tempAssetDir = path.join(tempAssetRootDir, entryId);
  const finalAssetDir = path.join(assetRootDir, entryId);
  let imageIndex = 0;

  await removePath(tempAssetDir);
  await ensureDir(tempAssetDir);

  try {
    const result = await mammoth.convertToHtml(
      { path: sourceFilePath },
      {
        convertImage: mammoth.images.imgElement(async (image) => {
          imageIndex += 1;
          const extension = getExtensionFromContentType(image.contentType);
          const fileName = `image-${imageIndex}.${extension}`;
          const base64 = await image.read('base64');
          const buffer = Buffer.from(base64, 'base64');

          await fsp.writeFile(path.join(tempAssetDir, fileName), buffer);

          return {
            src: `${PUBLIC_ASSET_PREFIX}/${entryId}/${fileName}`
          };
        })
      }
    );

    const markdownBody = normalizeMarkdown(turndownService.turndown(result.value || ''));

    if (imageIndex > 0) {
      await removePath(finalAssetDir);
      await ensureDir(path.dirname(finalAssetDir));
      await fsp.rename(tempAssetDir, finalAssetDir);
    } else {
      await removePath(tempAssetDir);
      await removePath(finalAssetDir);
    }

    return {
      markdown: buildMarkdownDocument({
        title,
        date: firstImportedAt,
        updated: updatedAt,
        body: markdownBody
      }),
      messages: result.messages || [],
      relativePath
    };
  } catch (error) {
    await removePath(tempAssetDir);
    throw error;
  }
}

async function cleanupRemovedEntries({ previousEntries, nextEntries, postDir, assetDir }) {
  for (const [relativePath, entry] of Object.entries(previousEntries)) {
    if (nextEntries[relativePath]) {
      continue;
    }

    const legacyPostFileName = `${entry.id}.md`;
    const currentPostFileName = entry.postFileName || legacyPostFileName;

    await removePath(path.join(postDir, legacyPostFileName));
    await removePath(path.join(postDir, currentPostFileName));
    await removePath(path.join(assetDir, entry.id));
  }
}

function createSyncController(hexo) {
  const baseDir = hexo.base_dir;
  const sourceDir = path.join(baseDir, SOURCE_FOLDER_NAME);
  const postDir = path.join(baseDir, MANAGED_POST_DIR);
  const assetDir = path.join(baseDir, MANAGED_ASSET_DIR);
  const manifestFilePath = path.join(baseDir, MANIFEST_PATH);
  const tempAssetRootDir = path.join(baseDir, '.cache', 'docx-essay-temp-assets');

  let syncQueue = Promise.resolve();
  let watcher = null;
  let debounceTimer = null;

  async function syncOnce(trigger) {
    const manifest = await readJson(manifestFilePath, {
      version: 1,
      sourceFolder: SOURCE_FOLDER_NAME,
      entries: {}
    });

    const previousEntries = manifest.entries || {};
    const nextEntries = {};
    const docxFiles = await listDocxFiles(sourceDir);
    let writtenCount = 0;

    await ensureDir(postDir);
    await ensureDir(assetDir);

    for (const filePath of docxFiles) {
      const relativePath = normalizePath(path.relative(sourceDir, filePath));
      const previousEntry = previousEntries[relativePath] || {};
      const stat = await fsp.stat(filePath);
      const title = path.basename(filePath, path.extname(filePath));
      const entryId = previousEntry.id || buildEntryId(relativePath);
      const firstImportedAt = previousEntry.firstImportedAt || formatDate(new Date());
      const updatedAt = formatDate(stat.mtime);
      const postFileName = buildPostFileName(title, entryId);

      const result = await convertDocxToManagedMarkdown({
        sourceFilePath: filePath,
        relativePath,
        entryId,
        title,
        firstImportedAt,
        updatedAt,
        assetRootDir: assetDir,
        tempAssetRootDir
      });

      const legacyPostFilePath = path.join(postDir, `${entryId}.md`);
      const postFilePath = path.join(postDir, postFileName);
      const didWrite = await writeFileIfChanged(postFilePath, result.markdown);

      if (postFilePath !== legacyPostFilePath) {
        await removePath(legacyPostFilePath);
      }

      if (didWrite) {
        writtenCount += 1;
      }

      for (const message of result.messages) {
        hexo.log.warn(`[docx-sync] ${relativePath}: ${message.message}`);
      }

      nextEntries[relativePath] = {
        id: entryId,
        title,
        postFileName,
        firstImportedAt
      };
    }

    await cleanupRemovedEntries({
      previousEntries,
      nextEntries,
      postDir,
      assetDir
    });

    await writeJsonIfChanged(manifestFilePath, {
      version: 1,
      sourceFolder: SOURCE_FOLDER_NAME,
      tag: TAG_NAME,
      entries: nextEntries
    });

    hexo.log.info(`[docx-sync] ${trigger}: synced ${docxFiles.length} docx file(s), updated ${writtenCount} post(s).`);
  }

  function enqueueSync(trigger) {
    syncQueue = syncQueue.then(() => syncOnce(trigger));
    return syncQueue;
  }

  function scheduleSync(trigger) {
    if (debounceTimer) {
      clearTimeout(debounceTimer);
    }

    debounceTimer = setTimeout(() => {
      enqueueSync(trigger).catch((error) => {
        hexo.log.error('[docx-sync] Sync failed while watching external docx folder.');
        hexo.log.error(error);
      });
    }, 300);
  }

  function startWatcher() {
    if (watcher) {
      return;
    }

    watcher = chokidar.watch(sourceDir, {
      ignoreInitial: true,
      awaitWriteFinish: {
        stabilityThreshold: 500,
        pollInterval: 100
      }
    });

    watcher
      .on('add', (filePath) => {
        if (/\.docx$/i.test(filePath)) {
          hexo.log.info(`[docx-sync] detected add: ${normalizePath(path.relative(sourceDir, filePath))}`);
          scheduleSync(`watch:add:${path.basename(filePath)}`);
        }
      })
      .on('change', (filePath) => {
        if (/\.docx$/i.test(filePath)) {
          hexo.log.info(`[docx-sync] detected change: ${normalizePath(path.relative(sourceDir, filePath))}`);
          scheduleSync(`watch:change:${path.basename(filePath)}`);
        }
      })
      .on('unlink', (filePath) => {
        if (/\.docx$/i.test(filePath)) {
          hexo.log.info(`[docx-sync] detected unlink: ${normalizePath(path.relative(sourceDir, filePath))}`);
          scheduleSync(`watch:unlink:${path.basename(filePath)}`);
        }
      })
      .on('error', (error) => {
        hexo.log.error('[docx-sync] Watcher error.');
        hexo.log.error(error);
      });

    hexo.log.info(`[docx-sync] Watching ${SOURCE_FOLDER_NAME} for docx changes.`);
  }

  return {
    enqueueSync,
    startWatcher
  };
}

const controller = createSyncController(hexo);

hexo.extend.filter.register('after_init', async function () {
  await controller.enqueueSync('after_init');

  if (this.env.cmd === 'server') {
    controller.startWatcher();
  }
});
