import { db } from "./firebase";
import { doc, setDoc, onSnapshot } from "firebase/firestore";

const COLLECTION = "household-ledger";

// Subscribes to a document and calls onData(parsedValue | null) whenever it
// changes — including changes made by the other person, live, no refresh needed.
export function subscribeDoc(key, onData) {
  const ref = doc(db, COLLECTION, key);
  return onSnapshot(
    ref,
    (snap) => {
      if (snap.exists() && snap.data() && snap.data().value !== undefined) {
        try {
          onData(JSON.parse(snap.data().value));
        } catch (e) {
          console.error("Failed to parse stored value for", key, e);
          onData(null);
        }
      } else {
        onData(null);
      }
    },
    (err) => {
      console.error("Firestore read error for", key, err);
      onData(null);
    }
  );
}

export async function saveDoc(key, value) {
  try {
    await setDoc(doc(db, COLLECTION, key), { value: JSON.stringify(value), updatedAt: Date.now() });
    return true;
  } catch (e) {
    console.error("Firestore write error for", key, e);
    return false;
  }
}
