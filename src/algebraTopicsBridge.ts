import {
  GoogleAuthProvider,
  onAuthStateChanged,
  signInWithPopup,
  type User,
} from 'firebase/auth';
import {
  collection,
  deleteDoc,
  doc,
  getDoc,
  getDocs,
  query,
  serverTimestamp,
  setDoc,
  updateDoc,
  where,
} from 'firebase/firestore';
import { auth, db, handleFirestoreError, OperationType } from './firebaseSaveBridge';

export interface ReviewQuestion {
  id: string;
  prompt: string;
  choices: string[];
  correctIndex: number;
  explanation: string;
  standardCode: string;
  difficulty: 'easy' | 'medium' | 'hard';
}

export interface ReviewTopic {
  topicId: string;
  ownerUid?: string;
  unitOrder: number;
  unitName: string;
  topicTitle: string;
  studentSummary: string;
  standards: string[];
  questions: ReviewQuestion[];
  isPublished: boolean;
  isStarter?: boolean;
}

interface ExtractedLink {
  url: string;
  label: string;
  kind: 'doc' | 'slides' | 'sheet' | 'drive' | 'web';
  fileId?: string;
}

interface SheetRow {
  rowIndex: number;
  cells: Array<{ text: string; links: ExtractedLink[] }>;
}

interface DraftTopic {
  topicId: string;
  unitOrder: number;
  unitName: string;
  topicTitle: string;
  studentSummary: string;
  standards: string[];
  rowNumbers: number[];
  resources: ExtractedLink[];
  questions: ReviewQuestion[];
  resourceReport?: Array<{ label: string; kind: string; status: string; detail: string }>;
  isPublished: boolean;
  isGenerating?: boolean;
  isPublishing?: boolean;
  genError?: string;
}

const WORKSPACE_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets.readonly',
  'https://www.googleapis.com/auth/documents.readonly',
  'https://www.googleapis.com/auth/presentations.readonly',
  'https://www.googleapis.com/auth/drive.readonly',
];

const STARTER_TOPICS: ReviewTopic[] = [
  {
    topicId: 'starter_linear_equations',
    unitOrder: 1,
    unitName: 'Unit 1: Solving Equations',
    topicTitle: 'Solving Multi-Step Linear Equations',
    studentSummary: 'Combine like terms, distribute, and solve equations with variables on both sides.',
    standards: ['HSA.REI.B.3'],
    isPublished: true,
    isStarter: true,
    questions: [
      {
        id: 'st_le_1',
        prompt: 'Solve for x:  3(x - 4) = 18',
        choices: ['x = 2', 'x = 6', 'x = 10', 'x = 22'],
        correctIndex: 2,
        explanation: 'Divide both sides by 3 to get x - 4 = 6, then add 4 to get x = 10.',
        standardCode: 'HSA.REI.B.3',
        difficulty: 'easy',
      },
      {
        id: 'st_le_2',
        prompt: 'Solve for x:  5x - 7 = 2x + 11',
        choices: ['x = 4/3', 'x = 6', 'x = 18', 'x = -6'],
        correctIndex: 1,
        explanation: 'Subtract 2x from both sides (3x - 7 = 11), add 7 (3x = 18), then divide by 3.',
        standardCode: 'HSA.REI.B.3',
        difficulty: 'medium',
      },
      {
        id: 'st_le_3',
        prompt: 'Solve for y:  -2(y + 5) + 4 = 12',
        choices: ['y = -9', 'y = -4', 'y = 1', 'y = 9'],
        correctIndex: 0,
        explanation: 'Distribute to get -2y - 10 + 4 = 12, so -2y - 6 = 12, -2y = 18, y = -9.',
        standardCode: 'HSA.REI.B.3',
        difficulty: 'medium',
      },
      {
        id: 'st_le_4',
        prompt: 'How many solutions does  4(x + 2) = 4x + 8  have?',
        choices: ['No solution', 'Exactly one solution (x = 2)', 'Exactly one solution (x = 0)', 'Infinitely many solutions'],
        correctIndex: 3,
        explanation: 'Distributing gives 4x + 8 = 4x + 8, which is an identity true for every x.',
        standardCode: 'HSA.REI.B.3',
        difficulty: 'medium',
      },
    ],
  },
  {
    topicId: 'starter_linear_functions',
    unitOrder: 2,
    unitName: 'Unit 2: Linear Functions',
    topicTitle: 'Slope, Intercepts & Writing Linear Equations',
    studentSummary: 'Find slope from two points and write equations in slope-intercept form.',
    standards: ['HSF.IF.C.7a', 'HSA.CED.A.2'],
    isPublished: true,
    isStarter: true,
    questions: [
      {
        id: 'st_lf_1',
        prompt: 'What is the slope of the line passing through (2, -3) and (6, 9)?',
        choices: ['1/3', '3', '-3', '2'],
        correctIndex: 1,
        explanation: 'Use (y2 - y1)/(x2 - x1) = (9 - (-3)) / (6 - 2) = 12 / 4 = 3.',
        standardCode: 'HSF.IF.C.7a',
        difficulty: 'easy',
      },
      {
        id: 'st_lf_2',
        prompt: 'Write the equation of a line with slope -2 that passes through (0, 5).',
        choices: ['y = 5x - 2', 'y = -2x + 5', 'y = -2x - 5', 'y = 2x + 5'],
        correctIndex: 1,
        explanation: 'Since (0, 5) is the y-intercept, b = 5 and m = -2 gives y = -2x + 5.',
        standardCode: 'HSA.CED.A.2',
        difficulty: 'easy',
      },
      {
        id: 'st_lf_3',
        prompt: 'What is the x-intercept of the line  3x - 4y = 24 ?',
        choices: ['(8, 0)', '(0, -6)', '(-6, 0)', '(0, 8)'],
        correctIndex: 0,
        explanation: 'Set y = 0: 3x = 24, so x = 8, giving the x-intercept (8, 0).',
        standardCode: 'HSF.IF.C.7a',
        difficulty: 'medium',
      },
      {
        id: 'st_lf_4',
        prompt: 'Which line is parallel to  y = (2/3)x - 7  and passes through (3, 4)?',
        choices: ['y = -(3/2)x + 2', 'y = (2/3)x + 2', 'y = (2/3)x + 4', 'y = (3/2)x - 1/2'],
        correctIndex: 1,
        explanation: 'Parallel lines share slope 2/3. Plugging in (3, 4): 4 = (2/3)(3) + b → 4 = 2 + b → b = 2.',
        standardCode: 'HSA.CED.A.2',
        difficulty: 'hard',
      },
    ],
  },
  {
    topicId: 'starter_systems',
    unitOrder: 3,
    unitName: 'Unit 3: Systems of Equations',
    topicTitle: 'Solving Systems by Substitution & Elimination',
    studentSummary: 'Find the intersection point (x, y) of two linear equations.',
    standards: ['HSA.REI.C.6'],
    isPublished: true,
    isStarter: true,
    questions: [
      {
        id: 'st_sys_1',
        prompt: 'Solve the system:  x + y = 10  and  x - y = 4',
        choices: ['(7, 3)', '(3, 7)', '(6, 4)', '(8, 2)'],
        correctIndex: 0,
        explanation: 'Add the two equations: 2x = 14 so x = 7, then 7 + y = 10 gives y = 3.',
        standardCode: 'HSA.REI.C.6',
        difficulty: 'easy',
      },
      {
        id: 'st_sys_2',
        prompt: 'Solve the system:  y = 2x - 1  and  3x + y = 14',
        choices: ['(3, 5)', '(5, 3)', '(4, 7)', '(2, 3)'],
        correctIndex: 0,
        explanation: 'Substitute y = 2x - 1 into 3x + y = 14: 5x - 1 = 14 → x = 3, y = 2(3) - 1 = 5.',
        standardCode: 'HSA.REI.C.6',
        difficulty: 'medium',
      },
      {
        id: 'st_sys_3',
        prompt: 'Two lines have the same slope but different y-intercepts. How many solutions does the system have?',
        choices: ['One solution', 'Two solutions', 'Infinitely many solutions', 'No solution'],
        correctIndex: 3,
        explanation: 'Parallel lines with different y-intercepts never intersect, so there is no solution.',
        standardCode: 'HSA.REI.C.6',
        difficulty: 'easy',
      },
    ],
  },
  {
    topicId: 'starter_quadratics',
    unitOrder: 4,
    unitName: 'Unit 4: Quadratics & Factoring',
    topicTitle: 'Factoring & Solving Quadratic Equations',
    studentSummary: 'Factor trinomials, difference of squares, and solve quadratic equations.',
    standards: ['HSA.SSE.B.3a', 'HSA.REI.B.4b'],
    isPublished: true,
    isStarter: true,
    questions: [
      {
        id: 'st_q_1',
        prompt: 'Factor completely:  x² - 9x + 14',
        choices: ['(x - 2)(x - 7)', '(x + 2)(x + 7)', '(x - 1)(x - 14)', '(x + 2)(x - 7)'],
        correctIndex: 0,
        explanation: 'Find two numbers multiplying to +14 and adding to -9: -2 and -7.',
        standardCode: 'HSA.SSE.B.3a',
        difficulty: 'easy',
      },
      {
        id: 'st_q_2',
        prompt: 'What are the solutions to  x² - 25 = 0 ?',
        choices: ['x = 5 only', 'x = -5 and x = 5', 'x = 25 and x = -25', 'x = 0 and x = 5'],
        correctIndex: 1,
        explanation: 'Difference of squares factors as (x - 5)(x + 5) = 0, giving x = 5 or x = -5.',
        standardCode: 'HSA.REI.B.4b',
        difficulty: 'easy',
      },
      {
        id: 'st_q_3',
        prompt: 'Solve for x:  2x² + 5x - 3 = 0',
        choices: ['x = -3 and x = 1/2', 'x = 3 and x = -1/2', 'x = -3 and x = -1/2', 'x = 3/2 and x = -1'],
        correctIndex: 0,
        explanation: 'Factor as (2x - 1)(x + 3) = 0, so x = 1/2 or x = -3.',
        standardCode: 'HSA.REI.B.4b',
        difficulty: 'hard',
      },
    ],
  },
];

// Strictly in-memory OAuth access token cache (never stored in localStorage/sessionStorage)
let cachedWorkspaceToken: string | null = null;
let currentAuthUser: User | null = null;

// Sync state for the Teacher Sync Modal
let publishedTopicsCache: ReviewTopic[] = [];
let teacherDraftTopicsCache: ReviewTopic[] = [];
let syncStep: 'SHEET' | 'CLUSTER' | 'GENERATE' | 'MANAGE' = 'SHEET';
let recentSheets: Array<{ id: string; name: string; modifiedTime?: string }> = [];
let sheetInspection: {
  spreadsheetId: string;
  title: string;
  tabs: Array<{ sheetId: number; title: string }>;
  activeTab: string;
  headers: string[];
  rows: SheetRow[];
  standardCol: number;
  topicHintCol: number;
  resourceCols: number[];
} | null = null;
let draftTopics: DraftTopic[] = [];
let syncBusyMsg = '';
let syncErrorMsg = '';
let editingDraftIdx: number | null = null;

const listeners = new Set<() => void>();
function notifyListeners() {
  listeners.forEach((fn) => {
    try {
      fn();
    } catch {
      // Ignore listener errors
    }
  });
}

onAuthStateChanged(auth, (user) => {
  if (!user || !user.emailVerified) {
    cachedWorkspaceToken = null;
    currentAuthUser = null;
    teacherDraftTopicsCache = [];
  } else {
    if (currentAuthUser && currentAuthUser.uid !== user.uid) {
      cachedWorkspaceToken = null;
    }
    currentAuthUser = user;
    void loadTeacherOwnTopics();
  }
  renderTeacherModalIfOpen();
  notifyListeners();
});

export async function signInTeacherWorkspace(): Promise<string | null> {
  const provider = new GoogleAuthProvider();
  provider.setCustomParameters({ prompt: 'select_account' });
  for (const scope of WORKSPACE_SCOPES) {
    provider.addScope(scope);
  }
  const result = await signInWithPopup(auth, provider);
  const credential = GoogleAuthProvider.credentialFromResult(result);
  if (!credential?.accessToken) {
    throw new Error('Could not obtain Google Workspace access token. Please allow requested permissions.');
  }
  cachedWorkspaceToken = credential.accessToken;
  currentAuthUser = result.user;
  await loadTeacherOwnTopics();
  notifyListeners();
  return cachedWorkspaceToken;
}

export async function loadPublishedTopics(): Promise<ReviewTopic[]> {
  try {
    const q = query(collection(db, 'reviewTopics'), where('isPublished', '==', true));
    const snap = await getDocs(q);
    const items: ReviewTopic[] = [];
    snap.forEach((d) => {
      const data = d.data();
      if (data && typeof data.topicTitle === 'string' && Array.isArray(data.questions)) {
        items.push({
          topicId: String(data.topicId || d.id),
          ownerUid: data.ownerUid ? String(data.ownerUid) : undefined,
          unitOrder: Number(data.unitOrder) || 1,
          unitName: String(data.unitName || 'Algebra'),
          topicTitle: String(data.topicTitle),
          studentSummary: String(data.studentSummary || ''),
          standards: Array.isArray(data.standards) ? data.standards.map(String) : [],
          questions: data.questions as ReviewQuestion[],
          isPublished: Boolean(data.isPublished),
        });
      }
    });
    items.sort((a, b) => a.unitOrder - b.unitOrder || a.topicTitle.localeCompare(b.topicTitle));
    publishedTopicsCache = items;
    notifyListeners();
    return items;
  } catch (err) {
    console.warn('Could not load published Firestore topics, using starter topics:', err);
    return publishedTopicsCache;
  }
}

async function loadTeacherOwnTopics(): Promise<void> {
  const user = auth.currentUser;
  if (!user || !user.emailVerified) return;
  const path = 'reviewTopics';
  try {
    const q = query(collection(db, path), where('ownerUid', '==', user.uid));
    const snap = await getDocs(q);
    const items: ReviewTopic[] = [];
    snap.forEach((d) => {
      const data = d.data();
      items.push({
        topicId: String(data.topicId || d.id),
        ownerUid: String(data.ownerUid || user.uid),
        unitOrder: Number(data.unitOrder) || 1,
        unitName: String(data.unitName || 'Algebra'),
        topicTitle: String(data.topicTitle || 'Topic'),
        studentSummary: String(data.studentSummary || ''),
        standards: Array.isArray(data.standards) ? data.standards.map(String) : [],
        questions: Array.isArray(data.questions) ? (data.questions as ReviewQuestion[]) : [],
        isPublished: Boolean(data.isPublished),
      });
    });
    items.sort((a, b) => a.unitOrder - b.unitOrder || a.topicTitle.localeCompare(b.topicTitle));
    teacherDraftTopicsCache = items;
  } catch (err) {
    // Non-teacher accounts won't have ownerUid topics; ignore empty/permission cleanly unless debugging
    console.debug('Teacher topics check:', err);
  }
}

function sanitizeTopicId(raw: string): string {
  const cleaned = raw.replace(/[^a-zA-Z0-9_\-]/g, '_').slice(0, 128);
  return cleaned.length >= 2 ? cleaned : `topic_${Date.now().toString(36)}`;
}

async function saveTopicToFirestore(topic: DraftTopic | ReviewTopic): Promise<void> {
  const user = auth.currentUser;
  if (!user || !user.emailVerified) {
    throw new Error('Please sign in with your verified Google teacher account first.');
  }
  const topicId = sanitizeTopicId(topic.topicId);
  const path = `reviewTopics/${topicId}`;
  const docRef = doc(db, 'reviewTopics', topicId);

  const safeStandards = (topic.standards.length ? topic.standards : [topic.topicTitle])
    .map((s) => String(s).trim().slice(0, 120))
    .filter(Boolean)
    .slice(0, 25);
  if (!safeStandards.length) safeStandards.push('Algebra Standard');

  const safeQuestions = (topic.questions || []).slice(0, 50).map((q, idx) => ({
    id: String(q.id || `q_${idx}`).slice(0, 80),
    prompt: String(q.prompt || '').trim().slice(0, 500),
    choices: (q.choices || ['A', 'B', 'C', 'D']).slice(0, 4).map((c) => String(c).slice(0, 180)),
    correctIndex: Math.max(0, Math.min(3, Math.trunc(Number(q.correctIndex) || 0))),
    explanation: String(q.explanation || 'Check each step.').trim().slice(0, 320),
    standardCode: String(q.standardCode || safeStandards[0]).trim().slice(0, 120),
    difficulty: q.difficulty === 'easy' || q.difficulty === 'hard' ? q.difficulty : 'medium',
  }));

  if (safeQuestions.length === 0) {
    throw new Error('Generate or add at least 1 question before saving.');
  }

  let existingSnap;
  try {
    existingSnap = await getDoc(docRef);
  } catch (err) {
    handleFirestoreError(err, OperationType.GET, path);
  }

  const basePayload = {
    unitOrder: Math.max(1, Math.min(99, Math.trunc(Number(topic.unitOrder) || 1))),
    unitName: String(topic.unitName || 'Algebra').trim().slice(0, 120) || 'Algebra',
    topicTitle: String(topic.topicTitle || 'Algebra Topic').trim().slice(0, 160) || 'Algebra Topic',
    studentSummary: String(topic.studentSummary || 'Review topic').trim().slice(0, 300) || 'Review topic',
    standards: safeStandards,
    questions: safeQuestions,
    isPublished: Boolean(topic.isPublished),
    updatedAt: serverTimestamp(),
  };

  if (!existingSnap.exists()) {
    try {
      await setDoc(docRef, {
        topicId,
        ownerUid: user.uid,
        ...basePayload,
        createdAt: serverTimestamp(),
      });
    } catch (err) {
      handleFirestoreError(err, OperationType.CREATE, path);
    }
  } else {
    try {
      await updateDoc(docRef, basePayload);
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, path);
    }
  }

  await Promise.all([loadPublishedTopics(), loadTeacherOwnTopics()]);
}

async function deleteTopicFromFirestore(topicId: string): Promise<void> {
  const path = `reviewTopics/${topicId}`;
  try {
    await deleteDoc(doc(db, 'reviewTopics', topicId));
  } catch (err) {
    handleFirestoreError(err, OperationType.DELETE, path);
  }
  await Promise.all([loadPublishedTopics(), loadTeacherOwnTopics()]);
}

function getAvailableTopics(): ReviewTopic[] {
  return publishedTopicsCache.length > 0 ? publishedTopicsCache : STARTER_TOPICS;
}

function escapeHtml(str: string): string {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function renderGoogleSignInButtonHtml(id: string, label: string): string {
  return `
    <button id="${id}" class="gsi-material-button" type="button">
      <div class="gsi-material-button-state"></div>
      <div class="gsi-material-button-content-wrapper">
        <div class="gsi-material-button-icon">
          <svg version="1.1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48" style="display: block; width: 20px; height: 20px;">
            <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"></path>
            <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"></path>
            <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"></path>
            <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"></path>
            <path fill="none" d="M0 0h48v48H0z"></path>
          </svg>
        </div>
        <span class="gsi-material-button-contents">${escapeHtml(label)}</span>
      </div>
    </button>
  `;
}

function renderTeacherModalIfOpen(): void {
  const modal = document.getElementById('overlay-teacher-sync');
  if (!modal || modal.classList.contains('hidden')) return;
  renderTeacherSyncModal();
}

function renderTeacherSyncModal(): void {
  const body = document.getElementById('teacher-sync-body');
  if (!body) return;

  const user = auth.currentUser;
  const hasToken = Boolean(cachedWorkspaceToken);

  const navTabs = `
    <div class="flex flex-wrap items-center justify-between gap-2 pb-3 mb-4 border-b-2 border-slate-700">
      <div class="flex flex-wrap gap-2">
        <button data-ts-step="SHEET" class="pixel-btn ${syncStep === 'SHEET' ? 'btn-gold' : 'btn-slate'} px-3 py-1.5 text-[10px]">1. SHEET &amp; COLUMNS</button>
        <button data-ts-step="CLUSTER" ${!sheetInspection ? 'disabled' : ''} class="pixel-btn ${syncStep === 'CLUSTER' ? 'btn-gold' : 'btn-slate'} px-3 py-1.5 text-[10px]">2. ORGANIZE TOPICS</button>
        <button data-ts-step="GENERATE" ${draftTopics.length === 0 ? 'disabled' : ''} class="pixel-btn ${syncStep === 'GENERATE' ? 'btn-gold' : 'btn-slate'} px-3 py-1.5 text-[10px]">3. GENERATE &amp; PUBLISH</button>
        <button data-ts-step="MANAGE" class="pixel-btn ${syncStep === 'MANAGE' ? 'btn-gold' : 'btn-slate'} px-3 py-1.5 text-[10px]">PUBLISHED (${publishedTopicsCache.length})</button>
      </div>
      <div class="text-[10px] text-slate-300">
        ${user ? `SIGNED IN: <span class="text-emerald-300">${escapeHtml(user.email || '')}</span>` : 'NOT SIGNED IN'}
        ${hasToken ? ' &bull; <span class="text-sky-300">WORKSPACE CONNECTED</span>' : ''}
      </div>
    </div>
  `;

  const alertBanner = `
    ${syncBusyMsg ? `<div class="mb-3 p-3 bg-sky-950 border-2 border-sky-400 text-sky-200 text-xs">${escapeHtml(syncBusyMsg)}</div>` : ''}
    ${syncErrorMsg ? `<div class="mb-3 p-3 bg-rose-950 border-2 border-rose-400 text-rose-200 text-xs">${escapeHtml(syncErrorMsg)}</div>` : ''}
  `;

  if (!user || !hasToken) {
    body.innerHTML = `
      ${navTabs}
      ${alertBanner}
      <div class="pixel-box bg-slate-900 p-5 text-center space-y-4">
        <div class="text-sm sm:text-base text-yellow-300">CONNECT GOOGLE WORKSPACE TO SYNC YOUR SHEET</div>
        <p class="text-xs text-slate-300 max-w-xl mx-auto leading-relaxed">
          Sign in with your teacher Google account to read your standards spreadsheet along with linked Google Docs, Google Slides, and Drive PDFs. Students never need Drive access—they play from the published question banks.
        </p>
        <div class="flex justify-center">
          ${renderGoogleSignInButtonHtml('btn-ts-connect-google', 'Sign in with Google')}
        </div>
        ${publishedTopicsCache.length > 0 ? `<div class="text-[10px] text-emerald-300">${publishedTopicsCache.length} published topic(s) currently live for students.</div>` : ''}
      </div>
    `;
    document.getElementById('btn-ts-connect-google')?.addEventListener('click', async () => {
      syncErrorMsg = '';
      syncBusyMsg = 'Connecting to Google Workspace...';
      renderTeacherSyncModal();
      try {
        await signInTeacherWorkspace();
        syncBusyMsg = '';
      } catch (err: any) {
        syncBusyMsg = '';
        syncErrorMsg = err?.message || 'Google Workspace connection failed.';
      }
      renderTeacherSyncModal();
    });
    wireNavTabClicks();
    return;
  }

  if (syncStep === 'MANAGE') {
    const allLive = teacherDraftTopicsCache.length > 0 ? teacherDraftTopicsCache : publishedTopicsCache;
    body.innerHTML = `
      ${navTabs}
      ${alertBanner}
      <div class="space-y-3">
        <div class="flex items-center justify-between">
          <div class="text-xs sm:text-sm text-yellow-300">LIVE TOPICS IN FIRESTORE (${allLive.length})</div>
          <button id="btn-ts-refresh-live" class="pixel-btn btn-blue px-3 py-1.5 text-[10px]">REFRESH LIST</button>
        </div>
        ${allLive.length === 0 ? `<div class="pixel-box bg-slate-900 p-4 text-xs text-slate-300">No teacher topics published yet. Students currently see the ${STARTER_TOPICS.length} built-in starter topics until you publish from your Sheet.</div>` : ''}
        <div class="space-y-2 max-h-[55vh] overflow-y-auto pr-1">
          ${allLive.map((t) => `
            <div class="pixel-box bg-slate-900 p-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div>
                <div class="text-[10px] text-sky-300">${escapeHtml(t.unitName)} &bull; ${t.questions.length} QUESTIONS &bull; ${t.isPublished ? '<span class="text-emerald-400">PUBLISHED</span>' : '<span class="text-amber-300">DRAFT</span>'}</div>
                <div class="text-xs sm:text-sm text-white">${escapeHtml(t.topicTitle)}</div>
                <div class="text-[10px] text-slate-400">Standards: ${escapeHtml(t.standards.join(', '))}</div>
              </div>
              <div class="flex items-center gap-2 shrink-0">
                <button data-toggle-pub="${escapeHtml(t.topicId)}" class="pixel-btn ${t.isPublished ? 'btn-slate' : 'btn-green'} px-2.5 py-1.5 text-[9px]">${t.isPublished ? 'UNPUBLISH' : 'PUBLISH'}</button>
                <button data-delete-topic="${escapeHtml(t.topicId)}" class="pixel-btn btn-red px-2.5 py-1.5 text-[9px]">DELETE</button>
              </div>
            </div>
          `).join('')}
        </div>
      </div>
    `;
    wireNavTabClicks();
    document.getElementById('btn-ts-refresh-live')?.addEventListener('click', async () => {
      syncBusyMsg = 'Refreshing Firestore topics...';
      renderTeacherSyncModal();
      await Promise.all([loadPublishedTopics(), loadTeacherOwnTopics()]);
      syncBusyMsg = '';
      renderTeacherSyncModal();
    });
    body.querySelectorAll<HTMLElement>('[data-toggle-pub]').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const id = btn.getAttribute('data-toggle-pub');
        const target = allLive.find((x) => x.topicId === id);
        if (!target) return;
        syncBusyMsg = `Updating "${target.topicTitle}"...`;
        syncErrorMsg = '';
        renderTeacherSyncModal();
        try {
          await saveTopicToFirestore({ ...target, isPublished: !target.isPublished });
        } catch (err: any) {
          syncErrorMsg = err?.message || 'Failed to update topic.';
        }
        syncBusyMsg = '';
        renderTeacherSyncModal();
      });
    });
    body.querySelectorAll<HTMLElement>('[data-delete-topic]').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const id = btn.getAttribute('data-delete-topic');
        if (!id) return;
        syncBusyMsg = 'Deleting topic...';
        syncErrorMsg = '';
        renderTeacherSyncModal();
        try {
          await deleteTopicFromFirestore(id);
        } catch (err: any) {
          syncErrorMsg = err?.message || 'Failed to delete topic.';
        }
        syncBusyMsg = '';
        renderTeacherSyncModal();
      });
    });
    return;
  }

  if (syncStep === 'SHEET') {
    body.innerHTML = `
      ${navTabs}
      ${alertBanner}
      <div class="space-y-4">
        <div class="pixel-box bg-slate-900 p-4 space-y-3">
          <div class="text-xs sm:text-sm text-yellow-300">PASTE YOUR GOOGLE SHEET LINK OR PICK FROM DRIVE</div>
          <div class="flex flex-col sm:flex-row gap-2">
            <input id="inp-sheet-url" type="text" placeholder="https://docs.google.com/spreadsheets/d/..." value="${escapeHtml(sheetInspection?.spreadsheetId || '')}" class="flex-1 bg-slate-950 border-2 border-slate-600 px-3 py-2 text-xs text-white focus:border-yellow-400 outline-none select-text" style="user-select:text;-webkit-user-select:text;" />
            <button id="btn-ts-inspect-sheet" class="pixel-btn btn-green px-4 py-2 text-xs shrink-0">LOAD SHEET</button>
            <button id="btn-ts-list-sheets" class="pixel-btn btn-blue px-3 py-2 text-xs shrink-0">MY RECENT SHEETS</button>
          </div>
          ${recentSheets.length > 0 ? `
            <div class="pt-2 border-t border-slate-800">
              <div class="text-[10px] text-slate-400 mb-1.5">RECENT GOOGLE SHEETS IN YOUR DRIVE:</div>
              <div class="flex flex-wrap gap-1.5 max-h-28 overflow-y-auto">
                ${recentSheets.map((f) => `<button data-pick-sheet="${escapeHtml(f.id)}" class="pixel-btn btn-slate px-2.5 py-1 text-[9px] normal-case">${escapeHtml(f.name)}</button>`).join('')}
              </div>
            </div>
          ` : ''}
        </div>

        ${sheetInspection ? `
          <div class="pixel-box bg-slate-900 p-4 space-y-3">
            <div class="flex flex-wrap items-center justify-between gap-2">
              <div>
                <div class="text-xs sm:text-sm text-emerald-300">LOADED: ${escapeHtml(sheetInspection.title)}</div>
                <div class="text-[10px] text-slate-400">${sheetInspection.rows.length} non-empty rows detected</div>
              </div>
              <div class="flex items-center gap-2">
                <label class="text-[10px] text-slate-300">TAB:</label>
                <select id="sel-sheet-tab" class="bg-slate-950 border-2 border-slate-600 px-2 py-1 text-xs text-white">
                  ${sheetInspection.tabs.map((t) => `<option value="${escapeHtml(t.title)}" ${t.title === sheetInspection!.activeTab ? 'selected' : ''}>${escapeHtml(t.title)}</option>`).join('')}
                </select>
              </div>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs">
              <div class="bg-slate-950 border border-slate-700 p-3 space-y-2">
                <label class="block text-[10px] text-yellow-300">STANDARDS COLUMN:</label>
                <select id="sel-std-col" class="w-full bg-slate-900 border border-slate-600 px-2 py-1.5 text-xs text-white">
                  ${sheetInspection.headers.map((h, i) => `<option value="${i}" ${i === sheetInspection!.standardCol ? 'selected' : ''}>${escapeHtml(h)}</option>`).join('')}
                </select>

                <label class="block text-[10px] text-sky-300 pt-1">OPTIONAL TOPIC / UNIT HINT COLUMN:</label>
                <select id="sel-hint-col" class="w-full bg-slate-900 border border-slate-600 px-2 py-1.5 text-xs text-white">
                  <option value="-1">(None — translate directly from Standards)</option>
                  ${sheetInspection.headers.map((h, i) => `<option value="${i}" ${i === sheetInspection!.topicHintCol ? 'selected' : ''}>${escapeHtml(h)}</option>`).join('')}
                </select>
              </div>

              <div class="bg-slate-950 border border-slate-700 p-3 space-y-2">
                <div class="text-[10px] text-emerald-300">RESOURCE LINK COLUMNS (DOCS / SLIDES / PDFS / ONLINE):</div>
                <div class="grid grid-cols-2 gap-1.5 max-h-32 overflow-y-auto">
                  ${sheetInspection.headers.map((h, i) => `
                    <label class="flex items-center gap-1.5 text-[10px] text-slate-200 cursor-pointer">
                      <input type="checkbox" class="res-col-chk" value="${i}" ${sheetInspection!.resourceCols.includes(i) ? 'checked' : ''} />
                      <span class="truncate">${escapeHtml(h)}</span>
                    </label>
                  `).join('')}
                </div>
              </div>
            </div>

            <div class="flex flex-wrap items-center justify-between gap-2 pt-2">
              <div class="text-[10px] text-slate-400">Preview: ${escapeHtml(sheetInspection.rows.slice(0, 3).map((r) => r.cells[sheetInspection!.standardCol]?.text || '').filter(Boolean).join(' | ').slice(0, 120))}</div>
              <button id="btn-ts-cluster" class="pixel-btn btn-gold px-4 py-2.5 text-xs">ORGANIZE &amp; DEDUPLICATE TOPICS &rarr;</button>
            </div>
          </div>
        ` : ''}
      </div>
    `;

    wireNavTabClicks();
    document.getElementById('btn-ts-list-sheets')?.addEventListener('click', async () => {
      syncErrorMsg = '';
      syncBusyMsg = 'Loading your recent Google Sheets...';
      renderTeacherSyncModal();
      try {
        const res = await fetch('/api/workspace/list-sheets', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${cachedWorkspaceToken}`,
          },
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || 'Failed to list sheets.');
        recentSheets = data.files || [];
      } catch (err: any) {
        syncErrorMsg = err?.message || 'Could not load recent spreadsheets.';
      }
      syncBusyMsg = '';
      renderTeacherSyncModal();
    });

    const loadSheetByValue = async (val: string, tabTitle?: string) => {
      syncErrorMsg = '';
      syncBusyMsg = 'Reading spreadsheet tabs, standards, and embedded hyperlinks...';
      renderTeacherSyncModal();
      try {
        const res = await fetch('/api/workspace/inspect-sheet', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${cachedWorkspaceToken}`,
          },
          body: JSON.stringify({ sheetUrlOrId: val, tabTitle }),
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || 'Failed to read Google Sheet.');
        sheetInspection = {
          spreadsheetId: data.spreadsheetId,
          title: data.title,
          tabs: data.tabs,
          activeTab: data.activeTab,
          headers: data.headers,
          rows: data.rows,
          standardCol: data.suggestedMapping.standardCol,
          topicHintCol: data.suggestedMapping.topicHintCol,
          resourceCols: data.suggestedMapping.resourceCols,
        };
      } catch (err: any) {
        syncErrorMsg = err?.message || 'Failed to inspect Google Sheet.';
      }
      syncBusyMsg = '';
      renderTeacherSyncModal();
    };

    document.getElementById('btn-ts-inspect-sheet')?.addEventListener('click', () => {
      const inp = document.getElementById('inp-sheet-url') as HTMLInputElement | null;
      if (inp?.value.trim()) void loadSheetByValue(inp.value.trim());
    });
    body.querySelectorAll<HTMLElement>('[data-pick-sheet]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = btn.getAttribute('data-pick-sheet');
        if (id) void loadSheetByValue(id);
      });
    });
    document.getElementById('sel-sheet-tab')?.addEventListener('change', (e) => {
      const val = (e.target as HTMLSelectElement).value;
      if (sheetInspection) void loadSheetByValue(sheetInspection.spreadsheetId, val);
    });
    document.getElementById('sel-std-col')?.addEventListener('change', (e) => {
      if (sheetInspection) sheetInspection.standardCol = Number((e.target as HTMLSelectElement).value);
    });
    document.getElementById('sel-hint-col')?.addEventListener('change', (e) => {
      if (sheetInspection) sheetInspection.topicHintCol = Number((e.target as HTMLSelectElement).value);
    });
    body.querySelectorAll<HTMLInputElement>('.res-col-chk').forEach((chk) => {
      chk.addEventListener('change', () => {
        if (!sheetInspection) return;
        const selected: number[] = [];
        body.querySelectorAll<HTMLInputElement>('.res-col-chk:checked').forEach((c) => selected.push(Number(c.value)));
        sheetInspection.resourceCols = selected;
      });
    });
    document.getElementById('btn-ts-cluster')?.addEventListener('click', async () => {
      if (!sheetInspection) return;
      syncErrorMsg = '';
      syncBusyMsg = 'Grouping all standards together into deduplicated, student-friendly Units & Topics...';
      renderTeacherSyncModal();
      try {
        const res = await fetch('/api/workspace/cluster-topics', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            sheetTitle: sheetInspection.title,
            headers: sheetInspection.headers,
            rows: sheetInspection.rows,
            standardCol: sheetInspection.standardCol,
            topicHintCol: sheetInspection.topicHintCol,
            resourceCols: sheetInspection.resourceCols,
          }),
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || 'Could not organize topics.');
        draftTopics = (data.topics || []).map((t: any) => ({
          ...t,
          questions: [],
          isPublished: true,
        }));
        syncStep = 'CLUSTER';
      } catch (err: any) {
        syncErrorMsg = err?.message || 'Clustering failed.';
      }
      syncBusyMsg = '';
      renderTeacherSyncModal();
    });
    return;
  }

  if (syncStep === 'CLUSTER') {
    body.innerHTML = `
      ${navTabs}
      ${alertBanner}
      <div class="space-y-3">
        <div class="flex flex-wrap items-center justify-between gap-2">
          <div class="text-xs text-slate-200">
            Review &amp; customize student-friendly topics below. Standards stay hidden behind the scenes for students.
          </div>
          <button id="btn-ts-go-generate" class="pixel-btn btn-green px-4 py-2 text-xs">NEXT: GENERATE QUESTIONS &rarr;</button>
        </div>

        <div class="space-y-2.5 max-h-[55vh] overflow-y-auto pr-1">
          ${draftTopics.map((t, idx) => `
            <div class="pixel-box bg-slate-900 p-3 space-y-2">
              <div class="grid grid-cols-1 sm:grid-cols-12 gap-2 items-center">
                <input data-draft-field="unitOrder" data-idx="${idx}" type="number" min="1" max="99" value="${t.unitOrder}" class="sm:col-span-1 bg-slate-950 border border-slate-600 px-2 py-1 text-xs text-yellow-300 select-text" style="user-select:text;-webkit-user-select:text;" title="Unit Order" />
                <input data-draft-field="unitName" data-idx="${idx}" type="text" value="${escapeHtml(t.unitName)}" class="sm:col-span-4 bg-slate-950 border border-slate-600 px-2 py-1 text-xs text-sky-300 select-text" style="user-select:text;-webkit-user-select:text;" placeholder="Unit Name" />
                <input data-draft-field="topicTitle" data-idx="${idx}" type="text" value="${escapeHtml(t.topicTitle)}" class="sm:col-span-5 bg-slate-950 border border-slate-600 px-2 py-1 text-xs text-white select-text" style="user-select:text;-webkit-user-select:text;" placeholder="Student Topic Title" />
                <div class="sm:col-span-2 flex justify-end gap-1">
                  ${idx > 0 ? `<button data-merge-up="${idx}" class="pixel-btn btn-blue px-2 py-1 text-[9px]" title="Merge into topic above">MERGE &uarr;</button>` : ''}
                  <button data-remove-draft="${idx}" class="pixel-btn btn-red px-2 py-1 text-[9px]">X</button>
                </div>
              </div>
              <div class="text-[10px] text-slate-300">
                Standards: <span class="text-amber-300">${escapeHtml(t.standards.join(' • '))}</span> &bull; Linked Resources: <span class="text-emerald-300">${t.resources.length}</span>
              </div>
            </div>
          `).join('')}
        </div>
      </div>
    `;
    wireNavTabClicks();
    body.querySelectorAll<HTMLInputElement>('[data-draft-field]').forEach((inp) => {
      inp.addEventListener('change', () => {
        const idx = Number(inp.getAttribute('data-idx'));
        const field = inp.getAttribute('data-draft-field') as 'unitOrder' | 'unitName' | 'topicTitle';
        if (!draftTopics[idx]) return;
        if (field === 'unitOrder') draftTopics[idx].unitOrder = Math.max(1, Math.min(99, Number(inp.value) || 1));
        else draftTopics[idx][field] = inp.value;
      });
    });
    body.querySelectorAll<HTMLElement>('[data-merge-up]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = Number(btn.getAttribute('data-merge-up'));
        if (idx <= 0 || !draftTopics[idx] || !draftTopics[idx - 1]) return;
        const target = draftTopics[idx - 1];
        const src = draftTopics[idx];
        target.standards = Array.from(new Set([...target.standards, ...src.standards])).slice(0, 25);
        const seenU = new Set(target.resources.map((r) => r.url));
        for (const r of src.resources) {
          if (!seenU.has(r.url)) {
            seenU.add(r.url);
            target.resources.push(r);
          }
        }
        target.rowNumbers = Array.from(new Set([...target.rowNumbers, ...src.rowNumbers]));
        draftTopics.splice(idx, 1);
        renderTeacherSyncModal();
      });
    });
    body.querySelectorAll<HTMLElement>('[data-remove-draft]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = Number(btn.getAttribute('data-remove-draft'));
        draftTopics.splice(idx, 1);
        renderTeacherSyncModal();
      });
    });
    document.getElementById('btn-ts-go-generate')?.addEventListener('click', () => {
      syncStep = 'GENERATE';
      renderTeacherSyncModal();
    });
    return;
  }

  // Step 3: GENERATE & PUBLISH
  const editingTopic = editingDraftIdx !== null ? draftTopics[editingDraftIdx] : null;
  body.innerHTML = `
    ${navTabs}
    ${alertBanner}
    <div class="space-y-3">
      <div class="flex flex-wrap items-center justify-between gap-2">
        <div class="text-xs text-slate-200">
          Generate questions from linked Docs/Slides/PDFs &amp; Standards, review, then publish to Firestore.
        </div>
        <div class="flex flex-wrap gap-2">
          <button id="btn-ts-gen-all" class="pixel-btn btn-blue px-3 py-1.5 text-[10px]">GENERATE ALL TOPICS</button>
          <button id="btn-ts-pub-all" class="pixel-btn btn-green px-3 py-1.5 text-[10px]">PUBLISH ALL READY</button>
        </div>
      </div>

      ${editingTopic ? `
        <div class="pixel-box-gold bg-slate-950 p-3 space-y-2">
          <div class="flex items-center justify-between">
            <div class="text-xs text-yellow-300">EDITING QUESTIONS: ${escapeHtml(editingTopic.topicTitle)} (${editingTopic.questions.length})</div>
            <button id="btn-ts-close-qedit" class="pixel-btn btn-slate px-2.5 py-1 text-[9px]">DONE EDITING</button>
          </div>
          <div class="space-y-2 max-h-[38vh] overflow-y-auto pr-1">
            ${editingTopic.questions.map((q, qIdx) => `
              <div class="border border-slate-700 bg-slate-900 p-2.5 space-y-1.5 text-[11px]">
                <div class="flex items-center justify-between gap-2">
                  <span class="text-sky-300">Q${qIdx + 1} (${escapeHtml(q.standardCode || '')})</span>
                  <button data-del-q="${qIdx}" class="pixel-btn btn-red px-2 py-0.5 text-[9px]">DELETE</button>
                </div>
                <input data-q-prompt="${qIdx}" value="${escapeHtml(q.prompt)}" class="w-full bg-slate-950 border border-slate-600 px-2 py-1 text-white select-text" style="user-select:text;-webkit-user-select:text;" />
                <div class="grid grid-cols-1 sm:grid-cols-2 gap-1.5">
                  ${q.choices.map((c, cIdx) => `
                    <div class="flex items-center gap-1">
                      <input type="radio" name="correct_${qIdx}" value="${cIdx}" ${q.correctIndex === cIdx ? 'checked' : ''} data-q-correct="${qIdx}" />
                      <input data-q-choice="${qIdx}_${cIdx}" value="${escapeHtml(c)}" class="flex-1 bg-slate-950 border border-slate-700 px-2 py-1 text-slate-200 select-text" style="user-select:text;-webkit-user-select:text;" />
                    </div>
                  `).join('')}
                </div>
              </div>
            `).join('')}
          </div>
        </div>
      ` : ''}

      <div class="space-y-2 max-h-[52vh] overflow-y-auto pr-1">
        ${draftTopics.map((t, idx) => `
          <div class="pixel-box bg-slate-900 p-3 flex flex-col gap-2">
            <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div>
                <div class="text-[10px] text-sky-300">${escapeHtml(t.unitName)} &bull; ${t.resources.length} resource(s) &bull; ${t.questions.length} question(s)</div>
                <div class="text-xs sm:text-sm text-white">${escapeHtml(t.topicTitle)}</div>
                ${t.genError ? `<div class="text-[10px] text-rose-400">${escapeHtml(t.genError)}</div>` : ''}
              </div>
              <div class="flex flex-wrap gap-1.5 shrink-0">
                <button data-gen-one="${idx}" ${t.isGenerating ? 'disabled' : ''} class="pixel-btn btn-blue px-2.5 py-1.5 text-[9px]">${t.isGenerating ? 'READING...' : (t.questions.length ? 'ADD MORE Qs' : 'GENERATE Qs')}</button>
                ${t.questions.length > 0 ? `<button data-edit-qs="${idx}" class="pixel-btn btn-slate px-2.5 py-1.5 text-[9px]">REVIEW (${t.questions.length})</button>` : ''}
                ${t.questions.length > 0 ? `<button data-pub-one="${idx}" ${t.isPublishing ? 'disabled' : ''} class="pixel-btn btn-green px-2.5 py-1.5 text-[9px]">${t.isPublishing ? 'SAVING...' : 'PUBLISH'}</button>` : ''}
              </div>
            </div>
            ${t.resourceReport && t.resourceReport.length ? `
              <div class="flex flex-wrap gap-1">
                ${t.resourceReport.map((rr) => `<span class="px-1.5 py-0.5 text-[9px] border ${rr.status === 'ok' ? 'border-emerald-500 text-emerald-300 bg-emerald-950/60' : 'border-amber-500 text-amber-300 bg-amber-950/60'}" title="${escapeHtml(rr.detail)}">${escapeHtml(rr.kind.toUpperCase())}: ${escapeHtml(rr.status === 'ok' ? 'OK' : 'SKIPPED')}</span>`).join('')}
              </div>
            ` : ''}
          </div>
        `).join('')}
      </div>
    </div>
  `;

  wireNavTabClicks();

  const generateForTopic = async (idx: number) => {
    const t = draftTopics[idx];
    if (!t) return;
    t.isGenerating = true;
    t.genError = '';
    renderTeacherSyncModal();
    try {
      const res = await fetch('/api/workspace/generate-topic-questions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${cachedWorkspaceToken}`,
        },
        body: JSON.stringify({
          topicTitle: t.topicTitle,
          unitName: t.unitName,
          studentSummary: t.studentSummary,
          standards: t.standards,
          resources: t.resources,
          existingPrompts: t.questions.map((q) => q.prompt),
          count: 8,
        }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Question generation failed.');
      t.questions = [...t.questions, ...(data.questions || [])].slice(0, 50);
      t.resourceReport = data.resourceReport || [];
    } catch (err: any) {
      t.genError = err?.message || 'Error generating questions.';
    } finally {
      t.isGenerating = false;
      renderTeacherSyncModal();
    }
  };

  body.querySelectorAll<HTMLElement>('[data-gen-one]').forEach((btn) => {
    btn.addEventListener('click', () => {
      void generateForTopic(Number(btn.getAttribute('data-gen-one')));
    });
  });

  document.getElementById('btn-ts-gen-all')?.addEventListener('click', async () => {
    for (let i = 0; i < draftTopics.length; i++) {
      if (draftTopics[i].questions.length === 0) {
        await generateForTopic(i);
      }
    }
  });

  body.querySelectorAll<HTMLElement>('[data-pub-one]').forEach((btn) => {
    btn.addEventListener('click', async () => {
      const idx = Number(btn.getAttribute('data-pub-one'));
      const t = draftTopics[idx];
      if (!t) return;
      t.isPublishing = true;
      syncErrorMsg = '';
      renderTeacherSyncModal();
      try {
        await saveTopicToFirestore(t);
        syncBusyMsg = `Published "${t.topicTitle}" to students!`;
      } catch (err: any) {
        syncErrorMsg = err?.message || 'Publish failed.';
      } finally {
        t.isPublishing = false;
        renderTeacherSyncModal();
      }
    });
  });

  document.getElementById('btn-ts-pub-all')?.addEventListener('click', async () => {
    const ready = draftTopics.filter((t) => t.questions.length > 0);
    if (!ready.length) {
      syncErrorMsg = 'Generate questions for at least one topic first.';
      renderTeacherSyncModal();
      return;
    }
    syncErrorMsg = '';
    syncBusyMsg = `Publishing ${ready.length} topic(s) to Firestore...`;
    renderTeacherSyncModal();
    try {
      for (const t of ready) {
        await saveTopicToFirestore(t);
      }
      syncBusyMsg = `Published ${ready.length} topic(s)! Students can now pick them in Quick Play & Dynasty.`;
    } catch (err: any) {
      syncErrorMsg = err?.message || 'Batch publish error.';
    }
    renderTeacherSyncModal();
  });

  body.querySelectorAll<HTMLElement>('[data-edit-qs]').forEach((btn) => {
    btn.addEventListener('click', () => {
      editingDraftIdx = Number(btn.getAttribute('data-edit-qs'));
      renderTeacherSyncModal();
    });
  });

  document.getElementById('btn-ts-close-qedit')?.addEventListener('click', () => {
    editingDraftIdx = null;
    renderTeacherSyncModal();
  });

  if (editingTopic) {
    body.querySelectorAll<HTMLElement>('[data-del-q]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const qIdx = Number(btn.getAttribute('data-del-q'));
        editingTopic.questions.splice(qIdx, 1);
        renderTeacherSyncModal();
      });
    });
    body.querySelectorAll<HTMLInputElement>('[data-q-prompt]').forEach((inp) => {
      inp.addEventListener('change', () => {
        const qIdx = Number(inp.getAttribute('data-q-prompt'));
        if (editingTopic.questions[qIdx]) editingTopic.questions[qIdx].prompt = inp.value;
      });
    });
    body.querySelectorAll<HTMLInputElement>('[data-q-choice]').forEach((inp) => {
      inp.addEventListener('change', () => {
        const [qIdxStr, cIdxStr] = (inp.getAttribute('data-q-choice') || '').split('_');
        const qIdx = Number(qIdxStr);
        const cIdx = Number(cIdxStr);
        if (editingTopic.questions[qIdx]) editingTopic.questions[qIdx].choices[cIdx] = inp.value;
      });
    });
    body.querySelectorAll<HTMLInputElement>('[data-q-correct]').forEach((inp) => {
      inp.addEventListener('change', () => {
        const qIdx = Number(inp.getAttribute('data-q-correct'));
        if (editingTopic.questions[qIdx]) editingTopic.questions[qIdx].correctIndex = Number(inp.value);
      });
    });
  }

  function wireNavTabClicks() {
    body?.querySelectorAll<HTMLElement>('[data-ts-step]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const next = btn.getAttribute('data-ts-step') as typeof syncStep;
        if (next) {
          syncStep = next;
          editingDraftIdx = null;
          renderTeacherSyncModal();
        }
      });
    });
  }
}

export function openTeacherSyncModal(): void {
  const modal = document.getElementById('overlay-teacher-sync');
  if (!modal) return;
  syncErrorMsg = '';
  syncBusyMsg = '';
  modal.classList.remove('hidden');
  renderTeacherSyncModal();
}

export function closeTeacherSyncModal(): void {
  const modal = document.getElementById('overlay-teacher-sync');
  if (modal) modal.classList.add('hidden');
}

declare global {
  interface Window {
    __algebraReviewBridge?: {
      getTopics: () => ReviewTopic[];
      refreshTopics: () => Promise<ReviewTopic[]>;
      openTeacherSync: () => void;
      closeTeacherSync: () => void;
      subscribe: (cb: () => void) => () => void;
    };
  }
}

window.__algebraReviewBridge = {
  getTopics: getAvailableTopics,
  refreshTopics: loadPublishedTopics,
  openTeacherSync: openTeacherSyncModal,
  closeTeacherSync: closeTeacherSyncModal,
  subscribe: (cb: () => void) => {
    listeners.add(cb);
    return () => listeners.delete(cb);
  },
};

void loadPublishedTopics();

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('btn-teacher-sync')?.addEventListener('click', openTeacherSyncModal);
  document.getElementById('btn-close-teacher-sync')?.addEventListener('click', closeTeacherSyncModal);
});
