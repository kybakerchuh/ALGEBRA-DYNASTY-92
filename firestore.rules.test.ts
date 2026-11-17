/**
 * Phase 0 TDD Security Rules Test Suite for Algebra Dynasty '92
 * Designed for @firebase/rules-unit-testing against firestore.rules.
 */
export interface DirtyDozenScenario {
  id: number;
  name: string;
  collectionPath: string;
  operation: 'create' | 'update' | 'get' | 'list' | 'delete';
  actorUid: string | null;
  actorEmail?: string | null;
  emailVerified: boolean;
  payload?: Record<string, unknown>;
  expectedOutcome: 'PERMISSION_DENIED';
}

export const validUserDoc = (uid: string) => ({
  uid,
  email: 'player@example.com',
  displayName: 'Coach 92',
  createdAt: 'REQUEST_TIME',
  updatedAt: 'REQUEST_TIME',
});

export const validCareerMap = () => ({
  teamId: 'BUF',
  season: 1,
  week: 1,
  phase: 'REGULAR',
  roster: [{ name: 'J. Kelly', pos: 'QB', ovr: 92 }],
  schedule: [{ week: 1, oppId: 'MIA', isHome: true, isBye: false, played: false }],
  standings: { BUF: { w: 0, l: 0, pf: 0, pa: 0 } },
});

export const validDeviceSaveDoc = (uid: string, deviceId = 'dev_abc123') => ({
  ownerUid: uid,
  deviceId,
  deviceLabel: 'Chrome Desktop',
  teamId: 'BUF',
  season: 1,
  week: 1,
  phase: 'REGULAR',
  rings: 0,
  careerData: validCareerMap(),
  createdAt: 'REQUEST_TIME',
  updatedAt: 'REQUEST_TIME',
});

export const validReviewTopicDoc = (ownerUid: string, topicId = 'topic_linear_1') => ({
  topicId,
  ownerUid,
  unitOrder: 1,
  unitName: 'Linear Equations',
  topicTitle: 'Solving Multi-Step Linear Equations',
  studentSummary: 'Solve equations with variables on both sides and distributive property.',
  standards: ['HSA.REI.B.3'],
  questions: [
    {
      id: 'q1',
      prompt: 'Solve for x: 3(x - 2) = 15',
      choices: ['x = 5', 'x = 7', 'x = 9', 'x = 3'],
      correctIndex: 1,
      explanation: 'Divide both sides by 3 to get x - 2 = 5, then add 2 to get x = 7.',
      standardCode: 'HSA.REI.B.3',
      difficulty: 'medium',
    },
  ],
  isPublished: true,
  createdAt: 'REQUEST_TIME',
  updatedAt: 'REQUEST_TIME',
});

export const DIRTY_DOZEN_PAYLOADS: DirtyDozenScenario[] = [
  {
    id: 1,
    name: 'Identity Spoofing on User Create',
    collectionPath: '/users/victim_uid',
    operation: 'create',
    actorUid: 'attacker_uid',
    emailVerified: true,
    payload: validUserDoc('victim_uid'),
    expectedOutcome: 'PERMISSION_DENIED',
  },
  {
    id: 2,
    name: 'Unverified Email Bypass on Teacher Topic Create',
    collectionPath: '/reviewTopics/topic_linear_1',
    operation: 'create',
    actorUid: 'spoof_uid',
    actorEmail: 'ky_baker@chuh.org',
    emailVerified: false,
    payload: validReviewTopicDoc('spoof_uid', 'topic_linear_1'),
    expectedOutcome: 'PERMISSION_DENIED',
  },
  {
    id: 3,
    name: 'Orphaned Subcollection Write (Missing Parent User Doc)',
    collectionPath: '/users/orphan_uid/deviceSaves/dev_1234',
    operation: 'create',
    actorUid: 'orphan_uid',
    emailVerified: true,
    payload: validDeviceSaveDoc('orphan_uid', 'dev_1234'),
    expectedOutcome: 'PERMISSION_DENIED',
  },
  {
    id: 4,
    name: 'Cross-Owner Subcollection Hijack',
    collectionPath: '/users/victim_uid/deviceSaves/dev_1234',
    operation: 'create',
    actorUid: 'attacker_uid',
    emailVerified: true,
    payload: validDeviceSaveDoc('attacker_uid', 'dev_1234'),
    expectedOutcome: 'PERMISSION_DENIED',
  },
  {
    id: 5,
    name: 'Shadow / Ghost Field Injection on ReviewTopic Create',
    collectionPath: '/reviewTopics/topic_linear_1',
    operation: 'create',
    actorUid: 'teacher_uid',
    actorEmail: 'ky_baker@chuh.org',
    emailVerified: true,
    payload: {
      ...validReviewTopicDoc('teacher_uid', 'topic_linear_1'),
      teacherEmail: 'ky_baker@chuh.org',
    },
    expectedOutcome: 'PERMISSION_DENIED',
  },
  {
    id: 6,
    name: 'Shadow / Ghost Field Injection on Update',
    collectionPath: '/users/owner_uid/deviceSaves/dev_1234',
    operation: 'update',
    actorUid: 'owner_uid',
    emailVerified: true,
    payload: {
      ...validDeviceSaveDoc('owner_uid', 'dev_1234'),
      ghostHack: 'injected',
    },
    expectedOutcome: 'PERMISSION_DENIED',
  },
  {
    id: 7,
    name: 'Immutable Field Mutation (topicId / createdAt)',
    collectionPath: '/reviewTopics/topic_linear_1',
    operation: 'update',
    actorUid: 'teacher_uid',
    actorEmail: 'ky_baker@chuh.org',
    emailVerified: true,
    payload: {
      ...validReviewTopicDoc('teacher_uid', 'topic_other_99'),
    },
    expectedOutcome: 'PERMISSION_DENIED',
  },
  {
    id: 8,
    name: 'Client Timestamp Forgery',
    collectionPath: '/users/owner_uid/deviceSaves/dev_1234',
    operation: 'create',
    actorUid: 'owner_uid',
    emailVerified: true,
    payload: {
      ...validDeviceSaveDoc('owner_uid', 'dev_1234'),
      updatedAt: 1700000000000,
    },
    expectedOutcome: 'PERMISSION_DENIED',
  },
  {
    id: 9,
    name: 'ID Poisoning / Oversized Topic ID',
    collectionPath: `/reviewTopics/${'a'.repeat(200)}`,
    operation: 'create',
    actorUid: 'teacher_uid',
    actorEmail: 'ky_baker@chuh.org',
    emailVerified: true,
    payload: validReviewTopicDoc('teacher_uid', 'a'.repeat(200)),
    expectedOutcome: 'PERMISSION_DENIED',
  },
  {
    id: 10,
    name: 'Payload Size Poisoning (Oversized Questions Array)',
    collectionPath: '/reviewTopics/topic_linear_1',
    operation: 'create',
    actorUid: 'teacher_uid',
    actorEmail: 'ky_baker@chuh.org',
    emailVerified: true,
    payload: {
      ...validReviewTopicDoc('teacher_uid', 'topic_linear_1'),
      questions: Array.from({ length: 60 }, (_, idx) => ({
        id: `q${idx}`,
        prompt: '2x = 4',
        choices: ['1', '2', '3', '4'],
        correctIndex: 1,
        explanation: 'Divide by 2.',
        standardCode: 'HSA.REI.B.3',
        difficulty: 'easy',
      })),
    },
    expectedOutcome: 'PERMISSION_DENIED',
  },
  {
    id: 11,
    name: 'Draft Scraping via Unfiltered ReviewTopics List Query',
    collectionPath: '/reviewTopics',
    operation: 'list',
    actorUid: null,
    emailVerified: false,
    expectedOutcome: 'PERMISSION_DENIED',
  },
  {
    id: 12,
    name: 'Unauthorized Student Topic Write / Admin Self-Promotion',
    collectionPath: '/admins/student_uid',
    operation: 'create',
    actorUid: 'student_uid',
    actorEmail: 'student@chuh.org',
    emailVerified: true,
    payload: {
      uid: 'student_uid',
      role: 'teacher',
      createdAt: 'REQUEST_TIME',
    },
    expectedOutcome: 'PERMISSION_DENIED',
  },
];
