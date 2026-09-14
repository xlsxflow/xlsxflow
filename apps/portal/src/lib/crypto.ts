// Production-grade HMAC-SHA256 signing for license generation
export async function generateLicenseSignature(orderId: string, hardwareHash: string): Promise<string> {
  const data = `${orderId}:${hardwareHash}`;
  const encoder = new TextEncoder();
  const dataBuffer = encoder.encode(data);
  
  // Use a secure secret key from the environment, fallback to a default only for local dev sandbox
  const secretString = process.env.LICENSE_SECRET_KEY || 'dev_sandbox_secret_key_change_in_prod';
  const secretBuffer = encoder.encode(secretString);

  // Import the secret into Web Crypto API for HMAC
  const key = await crypto.subtle.importKey(
    'raw',
    secretBuffer,
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );

  // Sign the data
  const signatureBuffer = await crypto.subtle.sign('HMAC', key, dataBuffer);
  
  // Convert ArrayBuffer to Hex String
  const signatureArray = Array.from(new Uint8Array(signatureBuffer));
  const signatureHex = signatureArray.map(b => b.toString(16).padStart(2, '0')).join('');
  
  return signatureHex;
}

export async function createLicenseJson(orderId: string, hardwareHash: string) {
  const signature = await generateLicenseSignature(orderId, hardwareHash);
  return {
    orderId,
    hardwareHash,
    signature,
    issuedAt: new Date().toISOString(),
    tier: "PRO",
    price: "$5 PPP"
  };
}
