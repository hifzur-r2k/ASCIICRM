import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { 
  initializeFirestore, 
  persistentLocalCache, 
  persistentMultipleTabManager 
} from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDllUvbEQs41D4RA2_2FQKMNh_A2cL_G4o",
  authDomain: "crm-fix.firebaseapp.com",
  projectId: "crm-fix",
  storageBucket: "crm-fix.appspot.com",
  messagingSenderId: "501429628278",
  appId: "1:501429628278:web:your_app_id"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);

// Multi-tab persistent offline cache
const db = initializeFirestore(app, {
  localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() })
});

export { auth, db };