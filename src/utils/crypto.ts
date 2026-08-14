// Helper: Convert ArrayBuffer to Base64
function bufferToBase64(buf: ArrayBuffer): string {
  const arr = new Uint8Array(buf);
  let binary = '';
  for (let i = 0; i < arr.byteLength; i++) {
    binary += String.fromCharCode(arr[i]);
  }
  return btoa(binary);
}

// Helper: Convert Base64 to ArrayBuffer
function base64ToBuffer(base64: string): ArrayBuffer {
  const binary = atob(base64);
  const buf = new ArrayBuffer(binary.length);
  const arr = new Uint8Array(buf);
  for (let i = 0; i < binary.length; i++) {
    arr[i] = binary.charCodeAt(i);
  }
  return buf;
}

// Derive a 256-bit AES-GCM key from password and salt using PBKDF2
async function getCryptoKey(password: string, salt: Uint8Array): Promise<CryptoKey> {
  const enc = new TextEncoder();
  const passwordKey = await window.crypto.subtle.importKey(
    'raw',
    enc.encode(password),
    { name: 'PBKDF2' },
    false,
    ['deriveKey']
  );
  
  return window.crypto.subtle.deriveKey(
    {
      name: 'PBKDF2',
      salt: salt as any,
      iterations: 100000,
      hash: 'SHA-256'
    },
    passwordKey,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );
}

// Set up the password for the first time by generating validation hash and salt
export async function setupPassword(password: string): Promise<{ hash: string; salt: string }> {
  const salt = window.crypto.getRandomValues(new Uint8Array(16));
  const key = await getCryptoKey(password, salt);
  const enc = new TextEncoder();
  const testData = enc.encode("verification_token");
  const iv = window.crypto.getRandomValues(new Uint8Array(12));
  const encrypted = await window.crypto.subtle.encrypt(
    { name: 'AES-GCM', iv: iv as any },
    key,
    testData
  );
  
  return {
    salt: bufferToBase64(salt.buffer),
    hash: JSON.stringify({
      iv: bufferToBase64(iv.buffer),
      ciphertext: bufferToBase64(encrypted)
    })
  };
}

// Verify entered password and return CryptoKey if correct
export async function verifyPassword(password: string, saltBase64: string, hashJson: string): Promise<CryptoKey | null> {
  try {
    const salt = new Uint8Array(base64ToBuffer(saltBase64));
    const key = await getCryptoKey(password, salt);
    
    const hash = JSON.parse(hashJson);
    const iv = new Uint8Array(base64ToBuffer(hash.iv));
    const ciphertext = base64ToBuffer(hash.ciphertext);
    
    const decrypted = await window.crypto.subtle.decrypt(
      { name: 'AES-GCM', iv: iv as any },
      key,
      ciphertext
    );
    
    const dec = new TextDecoder();
    const token = dec.decode(decrypted);
    if (token === "verification_token") {
      return key;
    }
  } catch (err) {
    console.error("Password verification failed:", err);
  }
  return null;
}

// Encrypt JSON payload using derived CryptoKey
export async function encryptPayload(payload: any, key: CryptoKey): Promise<string> {
  const jsonStr = JSON.stringify(payload);
  const enc = new TextEncoder();
  const data = enc.encode(jsonStr);
  const iv = window.crypto.getRandomValues(new Uint8Array(12));
  
  const ciphertext = await window.crypto.subtle.encrypt(
    { name: 'AES-GCM', iv: iv as any },
    key,
    data
  );
  
  return JSON.stringify({
    iv: bufferToBase64(iv.buffer),
    ciphertext: bufferToBase64(ciphertext)
  });
}

// Decrypt JSON payload using derived CryptoKey
export async function decryptPayload(encryptedStr: string, key: CryptoKey): Promise<any> {
  const data = JSON.parse(encryptedStr);
  const iv = new Uint8Array(base64ToBuffer(data.iv));
  const ciphertext = base64ToBuffer(data.ciphertext);
  
  const decrypted = await window.crypto.subtle.decrypt(
    { name: 'AES-GCM', iv: iv as any },
    key,
    ciphertext
  );
  
  const dec = new TextDecoder();
  const jsonStr = dec.decode(decrypted);
  return JSON.parse(jsonStr);
}

// Convert Blob to Base64 for encryption
export function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const dataUrl = reader.result as string;
      const base64 = dataUrl.split(',')[1];
      resolve(base64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

// Convert Base64 back to Blob
export function base64ToBlob(base64: string, type: string): Blob {
  const buf = base64ToBuffer(base64);
  return new Blob([buf], { type });
}
