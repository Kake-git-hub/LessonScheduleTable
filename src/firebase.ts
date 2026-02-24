import { initializeApp } from 'firebase/app'
import { getAuth, signInAnonymously, type User } from 'firebase/auth'
import {
  addDoc,
  collection,
  deleteDoc,
  doc,
  getDoc,
  getDocs,
  getFirestore,
  limit,
  onSnapshot,
  orderBy,
  query,
  setDoc,
  type Unsubscribe,
} from 'firebase/firestore'
import type { MasterData, SessionData } from './types'

const firebaseConfig = {
  apiKey: 'AIzaSyDPs5KUD9j-Oa3lwxEo4so_4tasLlSYI7Q',
  authDomain: 'lessonscheduletable.firebaseapp.com',
  projectId: 'lessonscheduletable',
  storageBucket: 'lessonscheduletable.firebasestorage.app',
  messagingSenderId: '390016573631',
  appId: '1:390016573631:web:82b7d38d9fb53a9938e60a',
  measurementId: 'G-DF1C35G0VJ',
}

const app = initializeApp(firebaseConfig)
const db = getFirestore(app)
const auth = getAuth(app)

// ---------- Path helpers ----------
// Data is scoped per classroom:
//   classrooms/{classroomId}                       — classroom metadata
//   classrooms/{classroomId}/master/default         — master data
//   classrooms/{classroomId}/sessions/{sessionId}   — session data

const classroomRef = (classroomId: string) => doc(db, 'classrooms', classroomId)
const classroomSessionsCol = (classroomId: string) => collection(db, 'classrooms', classroomId, 'sessions')
const classroomSessionRef = (classroomId: string, sessionId: string) => doc(db, 'classrooms', classroomId, 'sessions', sessionId)
const classroomMasterRef = (classroomId: string) => doc(db, 'classrooms', classroomId, 'master', 'default')

// ---------- Anonymous Auth ----------

let authReady: Promise<User | null> | null = null

export const initAuth = (): Promise<User | null> => {
  if (authReady) return authReady
  authReady = signInAnonymously(auth)
    .then((cred) => cred.user)
    .catch((e) => {
      console.warn('Anonymous auth failed:', e)
      return null
    })
  return authReady
}

export const waitForAuth = (): Promise<User | null> => authReady ?? Promise.resolve(null)

// ---------- Diagnostics ----------

export type FirestoreDiag = {
  authenticated: boolean
  uid: string
  canRead: boolean
  sessionFound: boolean
  teacherCount: number
  studentCount: number
  error?: string
}

export const diagnoseFirestore = async (classroomId: string, sessionId: string): Promise<FirestoreDiag> => {
  const result: FirestoreDiag = {
    authenticated: auth.currentUser != null,
    uid: auth.currentUser?.uid ?? '',
    canRead: false,
    sessionFound: false,
    teacherCount: 0,
    studentCount: 0,
  }
  try {
    const snap = await getDoc(classroomSessionRef(classroomId, sessionId))
    result.canRead = true
    result.sessionFound = snap.exists()
    if (snap.exists()) {
      const data = snap.data() as SessionData
      result.teacherCount = data.teachers?.length ?? 0
      result.studentCount = data.students?.length ?? 0
    }
  } catch (e) {
    result.error = e instanceof Error ? e.message : String(e)
  }
  return result
}

// ---------- Classroom CRUD ----------

export type ClassroomInfo = {
  id: string
  name: string
  createdAt: number
}

export const watchClassrooms = (
  callback: (items: ClassroomInfo[]) => void,
  onError?: (error: Error) => void,
): Unsubscribe => {
  const q = query(collection(db, 'classrooms'))
  return onSnapshot(
    q,
    (snapshot) => {
      const items = snapshot.docs
        .filter((d) => d.data().name)
        .map((d) => ({
          id: d.id,
          name: (d.data().name as string) ?? d.id,
          createdAt: (d.data().createdAt as number) ?? 0,
        }))
        .sort((a, b) => a.createdAt - b.createdAt)
      callback(items)
    },
    (error) => {
      if (onError) onError(error)
      else console.error('watchClassrooms error:', error)
    },
  )
}

export const createClassroom = async (id: string, name: string): Promise<void> => {
  await setDoc(classroomRef(id), { name, createdAt: Date.now() })
}

export const deleteClassroom = async (id: string): Promise<void> => {
  const sessionsSnap = await getDocs(classroomSessionsCol(id))
  for (const sDoc of sessionsSnap.docs) {
    await deleteDoc(sDoc.ref)
  }
  try { await deleteDoc(doc(db, 'classrooms', id, 'master', 'default')) } catch { /* ignore */ }
  await deleteDoc(classroomRef(id))
}

// ---------- Session CRUD (classroom-scoped) ----------

export type SessionListItem = {
  id: string
  name: string
  createdAt: number
  updatedAt: number
}

export const watchSessionsList = (
  classroomId: string,
  callback: (items: SessionListItem[]) => void,
  onError?: (error: Error) => void,
): Unsubscribe => {
  const q = query(classroomSessionsCol(classroomId), orderBy('settings.createdAt', 'desc'))
  return onSnapshot(
    q,
    (snapshot) => {
      const items = snapshot.docs.map((docSnap) => {
        const data = docSnap.data() as SessionData
        const createdAt = data.settings.createdAt ?? 0
        const updatedAt = data.settings.updatedAt ?? createdAt
        return { id: docSnap.id, name: data.settings.name, createdAt, updatedAt }
      })
      callback(items)
    },
    (error) => {
      if (onError) onError(error)
      else console.error('watchSessionsList error:', error)
    },
  )
}

export const watchSession = (
  classroomId: string,
  sessionId: string,
  callback: (value: SessionData | null) => void,
  onError?: (error: Error) => void,
): Unsubscribe =>
  onSnapshot(
    classroomSessionRef(classroomId, sessionId),
    (snapshot) => {
      callback(snapshot.exists() ? (snapshot.data() as SessionData) : null)
    },
    (error) => {
      if (onError) onError(error)
      else console.error('watchSession error:', error)
    },
  )

export const loadSession = async (classroomId: string, sessionId: string): Promise<SessionData | null> => {
  const snapshot = await getDoc(classroomSessionRef(classroomId, sessionId))
  return snapshot.exists() ? (snapshot.data() as SessionData) : null
}

export const saveSession = async (classroomId: string, sessionId: string, data: SessionData): Promise<void> => {
  const now = Date.now()
  const createdAt = data.settings.createdAt ?? now
  const next: SessionData = {
    ...data,
    settings: { ...data.settings, createdAt, updatedAt: now },
  }
  await setDoc(classroomSessionRef(classroomId, sessionId), next)
}

export const saveAndVerify = async (classroomId: string, sessionId: string, data: SessionData): Promise<boolean> => {
  await saveSession(classroomId, sessionId, data)
  const snap = await getDoc(classroomSessionRef(classroomId, sessionId))
  return snap.exists()
}

export const deleteSession = async (classroomId: string, sessionId: string): Promise<void> => {
  await deleteDoc(classroomSessionRef(classroomId, sessionId))
}

// ---------- Master Data (classroom-scoped) ----------

export const watchMasterData = (
  classroomId: string,
  callback: (data: MasterData | null) => void,
  onError?: (error: Error) => void,
): Unsubscribe =>
  onSnapshot(
    classroomMasterRef(classroomId),
    (snapshot) => {
      callback(snapshot.exists() ? (snapshot.data() as MasterData) : null)
    },
    (error) => {
      if (onError) onError(error)
      else console.error('watchMasterData error:', error)
    },
  )

export const loadMasterData = async (classroomId: string): Promise<MasterData | null> => {
  const snapshot = await getDoc(classroomMasterRef(classroomId))
  return snapshot.exists() ? (snapshot.data() as MasterData) : null
}

export const saveMasterData = async (classroomId: string, data: MasterData): Promise<void> => {
  await setDoc(classroomMasterRef(classroomId), data)
}

/** Load all session IDs for a classroom (one-shot). */
export const loadAllSessionIds = async (classroomId: string): Promise<string[]> => {
  const snapshot = await getDocs(classroomSessionsCol(classroomId))
  return snapshot.docs.map((d) => d.id)
}

// ---------- Auto-Backup ----------

const classroomBackupsCol = (classroomId: string) => collection(db, 'classrooms', classroomId, 'backups')

export type BackupMeta = {
  id: string
  createdAt: number
  trigger: 'auto' | 'manual'
  sessionCount: number
  hasMasterData: boolean
}

export type BackupData = {
  createdAt: number
  trigger: 'auto' | 'manual'
  masterData: MasterData | null
  sessions: Record<string, SessionData>
}

/** Save a full backup of the classroom (master + all sessions). */
export const createBackup = async (classroomId: string, trigger: 'auto' | 'manual'): Promise<string> => {
  const masterData = await loadMasterData(classroomId)
  const sessionIds = await loadAllSessionIds(classroomId)
  const sessionResults = await Promise.all(
    sessionIds.map((sid) => loadSession(classroomId, sid).then((d) => ({ id: sid, data: d }))),
  )
  const sessions = Object.fromEntries(
    sessionResults.filter((r) => r.data).map((r) => [r.id, r.data]),
  ) as Record<string, SessionData>

  const backup: BackupData = {
    createdAt: Date.now(),
    trigger,
    masterData,
    sessions,
  }
  const ref = await addDoc(classroomBackupsCol(classroomId), backup)
  return ref.id
}

/** List backups for a classroom (newest first). */
export const listBackups = async (classroomId: string): Promise<BackupMeta[]> => {
  const q = query(classroomBackupsCol(classroomId), orderBy('createdAt', 'desc'))
  const snapshot = await getDocs(q)
  return snapshot.docs.map((d) => {
    const data = d.data() as BackupData
    return {
      id: d.id,
      createdAt: data.createdAt,
      trigger: data.trigger ?? 'manual',
      sessionCount: Object.keys(data.sessions ?? {}).length,
      hasMasterData: !!data.masterData,
    }
  })
}

/** Load a specific backup (full data). */
export const loadBackup = async (classroomId: string, backupId: string): Promise<BackupData | null> => {
  const snapshot = await getDoc(doc(classroomBackupsCol(classroomId), backupId))
  return snapshot.exists() ? (snapshot.data() as BackupData) : null
}

/** Delete a specific backup. */
export const deleteBackup = async (classroomId: string, backupId: string): Promise<void> => {
  await deleteDoc(doc(classroomBackupsCol(classroomId), backupId))
}

/** Get the most recent backup timestamp (0 if none). */
export const getLastBackupTime = async (classroomId: string): Promise<number> => {
  const q = query(classroomBackupsCol(classroomId), orderBy('createdAt', 'desc'), limit(1))
  const snapshot = await getDocs(q)
  if (snapshot.empty) return 0
  return (snapshot.docs[0].data() as BackupData).createdAt
}

/** Keep only the most recent N backups, delete older ones. */
export const cleanupOldBackups = async (classroomId: string, maxCount: number): Promise<void> => {
  const q = query(classroomBackupsCol(classroomId), orderBy('createdAt', 'desc'))
  const snapshot = await getDocs(q)
  const toDelete = snapshot.docs.slice(maxCount)
  for (const d of toDelete) {
    await deleteDoc(d.ref)
  }
}

/** Restore a backup: overwrite master data and sessions. */
export const restoreBackup = async (classroomId: string, backup: BackupData): Promise<void> => {
  if (backup.masterData) {
    await setDoc(classroomMasterRef(classroomId), backup.masterData)
  }
  for (const [sid, sData] of Object.entries(backup.sessions)) {
    if (sData) {
      await setDoc(classroomSessionRef(classroomId, sid), sData)
    }
  }
}


