import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth";

// ⚠️ REPLACE THIS BLOCK WITH YOUR ACTUAL KEYS FROM FIREBASE
const firebaseConfig = {
  apiKey: "AIzaSyDllUvbEQs41D4RA2_2FQKMNh_A2cL_G4o",
  authDomain: "crm-fix.firebaseapp.com",
  projectId: "crm-fix",
  storageBucket: "crm-fix.firebasestorage.app",
  messagingSenderId: "501429628278",
  appId: "1:501429628278:web:3a2f5d36eda203aee45e57"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);

// Initialize Cloud Firestore and export it so our app can use it
export const db = getFirestore(app);
export const auth = getAuth(app);