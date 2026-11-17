# Security Specification: Algebra Dynasty '92 Cloud Save & Review Topics

## Phase 0: Data Invariants & Threat Model

### 1. Core Data Invariants
1. **Master Gate & Ownership**: A `DeviceDynastySave` document at `/users/{userId}/deviceSaves/{deviceId}` CANNOT exist without a verified parent `UserProfile` document at `/users/{userId}` whose `uid` equals `request.auth.uid`.
2. **Verified Identity & Anti-Spoofing**: All user profile and save reads/writes require `request.auth != null`, `request.auth.uid == userId`, and `request.auth.token.email_verified == true`.
3. **Device Isolation**: Every `DeviceDynastySave` must bind `ownerUid == request.auth.uid` and `deviceId == deviceId` (matching `^[a-zA-Z0-9_\-]+$`, length 4–128). Neither `ownerUid`, `deviceId`, nor `createdAt` may ever change after creation.
4. **PII Isolation**: User `email` and `displayName` live strictly inside `/users/{userId}` and are readable only by `isOwner(userId)` (`get` only—no `list` allowed on `/users`). `ReviewTopic` documents contain NO PII (never store teacher email or student PII).
5. **Teacher-Only Topic Publishing**: Only a verified teacher admin (`request.auth.token.email_verified == true && request.auth.token.email == 'ky_baker@chuh.org'` or matching `/admins/{uid}`) may create, update, or delete `/reviewTopics/{topicId}` documents, and `ownerUid` must equal `request.auth.uid` and remain immutable.
6. **Student Topic Read Isolation**: Students may only `get` or `list` `/reviewTopics` where `resource.data.isPublished == true` (enforced directly in the rule on `list`, with zero `get()`/`exists()` calls in `allow list`). Unpublished draft topics remain private to the teacher owner.
7. **Bounded Native Payloads**: `ReviewTopic.standards` is capped at `1..25` strings (`>= 1` char, `<= 120` chars) and `ReviewTopic.questions` is capped at `1..50` structured maps. `careerData` MUST be a native Firestore `map`.
8. **Temporal Integrity**: `createdAt == request.time` on `create`; `updatedAt == request.time` on both `create` and `update`; `createdAt` is immutable on `update`.

---

### 2. The "Dirty Dozen" Adversarial Payloads

1. **Identity Spoofing on User Create**: Unverified user or attacker tries creating `/users/victim_uid` with `uid: "victim_uid"`.
2. **Unverified Email Bypass**: Authenticated user with `email_verified: false` tries creating `/users/{uid}` or publishing `/reviewTopics/{topicId}` using a spoofed email claim.
3. **Orphaned Subcollection Write (Master Gate Bypass)**: User writes `/users/{uid}/deviceSaves/dev_123` before `/users/{uid}` exists.
4. **Cross-Owner Subcollection Hijack**: Attacker (`attacker_uid`) writes `/users/victim_uid/deviceSaves/dev_123` with `ownerUid: "attacker_uid"`.
5. **Shadow / Ghost Field Injection (Create)**: Owner includes undeclared field `isAdmin: true` inside `DeviceDynastySave` or `teacherEmail` inside `ReviewTopic`.
6. **Shadow / Ghost Field Injection (Update)**: Owner or teacher updates a document adding undeclared `extraPrivilege: true`.
7. **Immutable Field Mutation (`ownerUid` / `topicId` / `createdAt`)**: Owner tries changing `deviceId`, `topicId`, `ownerUid`, or backdating `createdAt` during an update.
8. **Client Timestamp Forgery**: Client sends a hardcoded numeric/past `updatedAt` instead of `request.time`.
9. **ID Poisoning / Path Exhaustion**: Attacker passes a 300-character or invalid-character `deviceId` or `topicId` (e.g., `topic/../admin` or `'a'.repeat(300)`).
10. **Payload Type / Size Poisoning (`careerData` as string or oversized `questions` array)**: Client passes `careerData` as a stringified JSON blob or injects `> 50` questions in `ReviewTopic`.
11. **PII Scraping / Draft Scraping via List Query**: Non-teacher queries `list` on `/users` or queries `list` on `/reviewTopics` without `where('isPublished', '==', true)` to read unpublished drafts.
12. **Unauthorized Student Topic Write / Admin Escalation**: Student or non-teacher user attempts to create/modify `/reviewTopics/{topicId}` or write to `/admins/{uid}`.
