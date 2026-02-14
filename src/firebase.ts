import { initializeApp } from 'firebase/app'
import {
  doc,
  getDoc,
  getFirestore,
  onSnapshot,
  setDoc,
  type Unsubscribe,
} from 'firebase/firestore'
import type { SessionData } from './types'

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
  await setDoc(sessionRef(sessionId), data)
}
