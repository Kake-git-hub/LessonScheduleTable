import { initializeApp } from 'firebase/app'
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

const sessionRef = (sessionId: string) => doc(db, 'sessions', sessionId)

export type SessionListItem = {
  id: string
  name: string
  createdAt: number
  updatedAt: number
}

export const watchSessionsList = (
  callback: (items: SessionListItem[]) => void,
): Unsubscribe => {
  const q = query(collection(db, 'sessions'), orderBy('settings.createdAt', 'desc'))
  return onSnapshot(q, (snapshot) => {
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
  })
}

export const watchSession = (
  sessionId: string,
  callback: (value: SessionData | null) => void,
): Unsubscribe =>
  onSnapshot(sessionRef(sessionId), (snapshot) => {
    callback(snapshot.exists() ? (snapshot.data() as SessionData) : null)
  })

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

export const deleteSession = async (sessionId: string): Promise<void> => {
  await deleteDoc(sessionRef(sessionId))
}

// --- Master Data ---

const masterRef = doc(db, 'master', 'default')

export const watchMasterData = (
  callback: (data: MasterData | null) => void,
): Unsubscribe =>
  onSnapshot(masterRef, (snapshot) => {
    callback(snapshot.exists() ? (snapshot.data() as MasterData) : null)
  })

export const loadMasterData = async (): Promise<MasterData | null> => {
  const snapshot = await getDoc(masterRef)
  return snapshot.exists() ? (snapshot.data() as MasterData) : null
}

export const saveMasterData = async (data: MasterData): Promise<void> => {
  await setDoc(masterRef, data)
}
