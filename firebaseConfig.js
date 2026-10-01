const firebaseConfig = {
  apiKey: process.env.FIREBASE_API_KEY || "YOUR_FIREBASE_API_KEY",
  authDomain: process.env.FIREBASE_AUTH_DOMAIN || "guptha-4cabc.firebaseapp.com",
  projectId: process.env.FIREBASE_PROJECT_ID || "guptha-4cabc",
  storageBucket: process.env.FIREBASE_STORAGE_BUCKET || "guptha-4cabc.firebasestorage.app",
  messagingSenderId: process.env.FIREBASE_MESSAGING_SENDER_ID || "842048416457",
  appId: process.env.FIREBASE_APP_ID || "1:842048416457:web:d535df3332c2397831fb18",
  measurementId: process.env.FIREBASE_MEASUREMENT_ID || "G-V2PHSLK4NH"
};

export default firebaseConfig;
