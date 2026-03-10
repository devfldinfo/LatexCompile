/**
 * ============================
 * Configuration
 * ============================
 */
const DRIVE_OUTPUT_FOLDER_ID = '1fS0rwF5ceTEpk5nMkg4J9HW1cDqxmBNd';
const DRIVE_INPUT_FOLDER_ID = '1y8qwfeOtjauaFiKNL-l3VoYUTKca6Q-a';

const CONFIG = (() => {
  const props = PropertiesService.getScriptProperties();
  return {
    GITHUB_TOKEN: props.getProperty('GITHUB_TOKEN'),
    GITHUB_OWNER: props.getProperty('GITHUB_OWNER'),
    GITHUB_REPO: props.getProperty('GITHUB_REPO'),
    GITHUB_BRANCH: 'main',
    INPUT_ROOT: 'input',
    POLL_INTERVAL_MS: 10000,
    MAX_POLLS: 30 // ~5 minutes
  };
})();

/**
 * ============================
 * Entry Point
 * ============================
 * Run this on a time trigger (every 1 minute)
 */
function processUnreadEmails2() {
  try {
    resetGithubInputFolder();

    const threads = GmailApp.search('is:unread in:inbox newer_than:5d');
    const uploads = [];
    const jobs = [];

    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        if (!msg.isUnread()) return;

        const attachments = msg.getAttachments();
        if (!attachments.length) return;

        const sender = extractEmail(msg.getFrom());
        const timestamp = Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          'ddMMyyyyHHmmss'
        );

        const folder = `${CONFIG.INPUT_ROOT}/${sender} ${timestamp}`;

        attachments.forEach(a => {
          uploads.push({
            path: `${folder}/${a.getName()}`,
            content: Utilities.newBlob(a.getBytes()).getDataAsString()
          });
        });

        jobs.push({ sender, folder, message: msg });
      });
    });

    if (!uploads.length) return;

    createSingleCommit(uploads);

    pollUntilCompleted(CONFIG.INPUT_ROOT);

    jobs.forEach(job => {
      sendPDFs(job);
      job.message.markRead(); // ✅ marked ONLY after success
    });

  } catch (err) {
    throw err;
  }
}

function processUnreadEmails() {
  try {
    resetGithubInputFolder();

    const threads = GmailApp.search('is:unread in:inbox newer_than:5d');
    const uploads = [];
    const jobs = [];

    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        if (!msg.isUnread()) return;

        const attachments = msg.getAttachments();
        if (!attachments.length) return;

        // ✅ Check if email contains at least one .tex attachment
        const hasTex = attachments.some(a =>
          /\.tex$/i.test(a.getName())
        );

        // 🚫 Skip email if no .tex files present
        if (!hasTex) return;

        const sender = extractEmail(msg.getFrom());
        const timestamp = Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          'ddMMyyyyHHmmss'
        );

        const folder = `${CONFIG.INPUT_ROOT}/${sender} ${timestamp}`;

        // ✅ Save ALL attachments
        attachments.forEach(a => {
          uploads.push({
            path: `${folder}/${a.getName()}`,
            content: Utilities.newBlob(a.getBytes()).getDataAsString()
          });
        });

        jobs.push({ sender, folder, message: msg });
      });
    });

    if (!uploads.length) return;

    createSingleCommit(uploads);

    pollUntilCompleted(CONFIG.INPUT_ROOT);

    jobs.forEach(job => {
      sendPDFs(job);
      job.message.markRead(); // ✅ only after success
    });

  } catch (err) {
    throw err;
  }
}


/**
 * ============================
 * Polling Logic
 * ============================
 */
function pollUntilCompleted() {
  for (let i = 0; i < CONFIG.MAX_POLLS; i++) {
    if (completedExists(CONFIG.INPUT_ROOT)) return;
    Utilities.sleep(CONFIG.POLL_INTERVAL_MS);
  }
  throw new Error(`Timeout waiting for completed.txt in ${job.folder}`);
}

function completedExists(folder) {
  try {
    const files = githubRequest('get', `/contents/${folder}`);
    return files.some(f => f.name === 'completed.txt');
  } catch {
    return false;
  }
}

/**
 * ============================
 * Email PDFs
 * ============================
 */
function sendPDFs(job) {
  const pdfs = listPDFs(job.folder);

  if (!pdfs.length) {
    throw new Error(`No PDFs found in ${job.folder}`);
  }

  const attachments = pdfs.map(name =>
    downloadFromGitHub(`${job.folder}/${name}`)
  );

  GmailApp.sendEmail(
    job.sender,
    'Your compiled PDF(s)',
    'See attached PDF files.',
    { attachments }
  );
}

/**
 * ============================
 * GitHub Helpers
 * ============================
 */
function createSingleCommit(files) {
  const baseSha = getBranchSha();
  const treeSha = createTree(baseSha, files);
  const commitSha = createCommit(treeSha, baseSha);
  updateBranch(commitSha);
}

function getBranchSha() {
  return githubRequest(
    'get',
    `/git/ref/heads/${CONFIG.GITHUB_BRANCH}`
  ).object.sha;
}

function createTree(baseSha, files) {
  return githubRequest('post', '/git/trees', {
    base_tree: baseSha,
    tree: files.map(f => ({
      path: f.path,
      mode: '100644',
      type: 'blob',
      content: f.content
    }))
  }).sha;
}

function createCommit(treeSha, parentSha) {
  return githubRequest('post', '/git/commits', {
    message: 'Upload LaTeX files',
    tree: treeSha,
    parents: [parentSha]
  }).sha;
}

function updateBranch(commitSha) {
  githubRequest('patch', `/git/refs/heads/${CONFIG.GITHUB_BRANCH}`, {
    sha: commitSha
  });
}

function listPDFs(folder) {
  return githubRequest('get', `/contents/${folder}`)
    .filter(f => f.name.endsWith('.pdf'))
    .map(f => f.name);
}

function downloadFromGitHub(path) {
  const res = githubRequest('get', `/contents/${path}`);
  return Utilities.newBlob(
    Utilities.base64Decode(res.content),
    'application/pdf',
    res.name
  );
}

function githubRequest(method, path, payload) {
  const options = {
    method,
    headers: {
      Authorization: `Bearer ${CONFIG.GITHUB_TOKEN}`,
      Accept: 'application/vnd.github.v3+json'
    },
    muteHttpExceptions: false
  };

  if (payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }

  const res = UrlFetchApp.fetch(
    `https://api.github.com/repos/${CONFIG.GITHUB_OWNER}/${CONFIG.GITHUB_REPO}${path}`,
    options
  );

  return JSON.parse(res.getContentText());
}

/**
 * ============================
 * Utilities
 * ============================
 */
function extractEmail(from) {
  const m = from.match(/<(.+?)>/);
  return m ? m[1] : from.trim();
}

/**
 * ============================
 * Error Logging
 * ============================
 */
//function onError(e) {
//  const SHEET_ID = '1Y5AHU332Ff0YzfowsFD2PVaQXu22yL4ll9KCLwkjG9s';
//  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
//  sheet.appendRow([
//    new Date(),
//    e && e.message ? e.message : JSON.stringify(e)
//  ]);
//}

/* ================= RESET INPUT ================= */
function resetGithubInputFolder() {
  const base = githubApiBase();
  const token = CONFIG.GITHUB_TOKEN
  const branch = CONFIG.GITHUB_BRANCH;

  // 1. Get HEAD commit
  const ref = JSON.parse(
    githubFetch(`${base}/git/refs/heads/${branch}`, token)
  );
  const commitSha = ref.object.sha;

  // 2. Get commit → tree SHA
  const commit = JSON.parse(
    githubFetch(`${base}/git/commits/${commitSha}`, token)
  );
  const treeSha = commit.tree.sha;

  // 3. Load FULL recursive tree
  const treeRes = JSON.parse(
    githubFetch(`${base}/git/trees/${treeSha}?recursive=1`, token)
  );

  // 4. REMOVE input tree AND all children
  const cleanedTree = treeRes.tree
    .filter(entry =>
      entry.path !== "input" &&
      !entry.path.startsWith("input/")
    )
    .map(entry => ({
      path: entry.path,
      mode: entry.mode,
      type: entry.type,
      sha: entry.sha
    }));

  // 5. Create NEW tree (no base_tree)
  const newTree = JSON.parse(
    githubFetch(
      `${base}/git/trees`,
      token,
      "post",
      { tree: cleanedTree }
    )
  );

  // 6. Commit
  const newCommit = JSON.parse(
    githubFetch(
      `${base}/git/commits`,
      token,
      "post",
      {
        message: "Fully remove /input after PDF sync",
        tree: newTree.sha,
        parents: [commitSha]
      }
    )
  );

  // 7. Move branch
  githubFetch(
    `${base}/git/refs/heads/${branch}`,
    token,
    "patch",
    { sha: newCommit.sha }
  );
}


function githubApiBase() {
  const p = PropertiesService.getScriptProperties();
  return `https://api.github.com/repos/${p.getProperty("GITHUB_OWNER")}/${p.getProperty("GITHUB_REPO")}`;
}

function githubFetch(url, token, method, payload) {
  const opt = {
    method: method || "get",
    headers: { Authorization: "token " + token },
    muteHttpExceptions: true
  };
  if (payload) {
    opt.contentType = "application/json";
    opt.payload = JSON.stringify(payload);
  }

  const r = UrlFetchApp.fetch(url, opt);
  if (r.getResponseCode() >= 300)
  {
   // Logger.log(r.getContentText());
    throw new Error(r.getContentText());
  
  }
  return r.getContentText();
}

/************************************** */
function processDriveTexFolder() {

  resetGithubInputFolder();

  const folder = DriveApp.getFolderById(DRIVE_INPUT_FOLDER_ID);
  const files = folder.getFiles();

  const uploads = [];
  const jobs = [];

  while (files.hasNext()) {
    const file = files.next();

    if (!/\.tex$/i.test(file.getName())) continue;

    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'ddMMyyyyHHmmss'
    );

    const jobFolder = `${CONFIG.INPUT_ROOT}/drive ${timestamp}`;

    uploads.push({
      path: `${jobFolder}/${file.getName()}`,
      content: file.getBlob().getDataAsString()
    });

    jobs.push({
      folder: jobFolder,
      sourceFile: file
    });
  }

  if (!uploads.length) return;

  createSingleCommit(uploads);

  pollUntilCompleted(CONFIG.INPUT_ROOT);

  jobs.forEach(job => {
    savePDFsToDrive(job);
  });

}

function savePDFsToDrive(job) {

  const pdfs = listPDFs(job.folder);

  if (!pdfs.length) {
    throw new Error(`No PDFs found in ${job.folder}`);
  }

  const outputFolder = DriveApp.getFolderById(DRIVE_OUTPUT_FOLDER_ID);

  pdfs.forEach(name => {

    const blob = downloadFromGitHub(`${job.folder}/${name}`);

    outputFolder.createFile(blob);

  });

}
