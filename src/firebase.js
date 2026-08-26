import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDllUvbEQs41D4RA2_2FQKMNh_A2cL_G4o",
  authDomain: "crm-fix.firebaseapp.com",
  projectId: "crm-fix",
  storageBucket: "crm-fix.appspot.com",
  messagingSenderId: "501429628278",
  appId: "1:501429628278:web:your_app_id"
};

// 1. Primary App (Handles your main login and offline database)
const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);

// Multi-tab persistent offline cache
export const db = initializeFirestore(app, {
  localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() })
});

// 2. Secondary App (Only used silently to create Supervisors without logging you out)
const secondaryApp = initializeApp(firebaseConfig, "SecondaryApp");
export const secondaryAuth = getAuth(secondaryApp);