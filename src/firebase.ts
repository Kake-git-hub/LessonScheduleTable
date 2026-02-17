import { initializeApp } from 'firebase/app'
import { getAuth, signInAnonymously, type User } from 'firebase/auth'
import {
  collection,
  deleteDoc,
  doc,
  getDoc,
  getFirestore,
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

const sessionRef = (sessionId: string) => doc(db, 'sessions', sessionId)

// ---------- Anonymous Auth ----------

let authReady: Promise<User | null> | null = null

/** Sign in anonymously. Call once at app start. */
export const initAuth = (): Promise<User | null> => {
  if (authReady) return authReady
  authReady = signInAnonymously(auth)
    .then((cred) => cred.user)
    .catch((e) => {
      console.warn('Anonymous auth failed (enable it in Firebase Console → Authentication → Sign-in method):', e)
      return null
    })
  return authReady
}

/** Wait until auth is settled. Returns current user or null. */
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

/** Run a full diagnostic: auth state, read access, and session existence. */
export const diagnoseFirestore = async (sessionId: string): Promise<FirestoreDiag> => {
  const result: FirestoreDiag = {
    authenticated: auth.currentUser != null,
    uid: auth.currentUser?.uid ?? '',
    canRead: false,
    sessionFound: false,
    teacherCount: 0,
    studentCount: 0,
  }
  try {
    const snap = await getDoc(sessionRef(sessionId))
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

// ---------- Session CRUD ----------

export type SessionListItem = {
  id: string
  name: string
  createdAt: number
  updatedAt: number
}

export const watchSessionsList = (
  callback: (items: SessionListItem[]) => void,
  onError?: (error: Error) => void,
): Unsubscribe => {
  const q = query(collection(db, 'sessions'), orderBy('settings.createdAt', 'desc'))
  return onSnapshot(
    q,
    (snapshot) => {
      const items = snapshot.docs.map((docSnap) => {
        const data = docSnap.data() as SessionData
        const createdAt = data.settings.createdAt ?? 0
        const updatedAt = data.settings.updatedAt ?? createdAt
        return {
          id: docSnap.id,
          name: data.settings.name,
          createdAt,
          updatedAt,
        }
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
  sessionId: string,
  callback: (value: SessionData | null) => void,
  onError?: (error: Error) => void,
): Unsubscribe =>
  onSnapshot(
    sessionRef(sessionId),
    (snapshot) => {
      callback(snapshot.exists() ? (snapshot.data() as SessionData) : null)
    },
    (error) => {
      if (onError) onError(error)
      else console.error('watchSession error:', error)
    },
  )

export const loadSession = async (sessionId: string): Promise<SessionData | null> => {
  const snapshot = await getDoc(sessionRef(sessionId))
  return snapshot.exists() ? (snapshot.data() as SessionData) : null
}

export const saveSession = async (sessionId: string, data: SessionData): Promise<void> => {
  const now = Date.now()
  const createdAt = data.settings.createdAt ?? now
  const next: SessionData = {
    ...data,
    settings: {
      ...data.settings,
      createdAt,
      updatedAt: now,
    },
  }
  await setDoc(sessionRef(sessionId), next)
}

/** Save and immediately verify the write was committed to the server. */
export const saveAndVerify = async (sessionId: string, data: SessionData): Promise<boolean> => {
  await saveSession(sessionId, data)
  // Read back to verify server-side persistence
  const snap = await getDoc(sessionRef(sessionId))
  return snap.exists()
}

export const deleteSession = async (sessionId: string): Promise<void> => {
  await deleteDoc(sessionRef(sessionId))
}

// --- Master Data ---

const masterRef = doc(db, 'master', 'default')

export const watchMasterData = (
  callback: (data: MasterData | null) => void,
  onError?: (error: Error) => void,
): Unsubscribe =>
  onSnapshot(
    masterRef,
    (snapshot) => {
      callback(snapshot.exists() ? (snapshot.data() as MasterData) : null)
    },
    (error) => {
      if (onError) onError(error)
      else console.error('watchMasterData error:', error)
    },
  )

export const loadMasterData = async (): Promise<MasterData | null> => {
  const snapshot = await getDoc(masterRef)
  return snapshot.exists() ? (snapshot.data() as MasterData) : null
}

export const saveMasterData = async (data: MasterData): Promise<void> => {
  await setDoc(masterRef, data)
}
