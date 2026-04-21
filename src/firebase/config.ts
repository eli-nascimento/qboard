import { initializeApp } from "firebase/app";
import { getAuth, GoogleAuthProvider } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyBpO8THVJTjqk5G0nOLQmF2gJr9eU_2UPw",
  authDomain: "qboard-98e09.firebaseapp.com",
  projectId: "qboard-98e09",
  storageBucket: "qboard-98e09.firebasestorage.app",
  messagingSenderId: "873702607887",
  appId: "1:873702607887:web:86e3de301f06f82d73a77a"
};

const app = initializeApp(firebaseConfig);

export const auth = getAuth(app);
export const googleProvider = new GoogleAuthProvider();
export const db = getFirestore(app);