// Uses standard Web Crypto to create a stable 64-character hash
export async function generateHardwareFingerprint(userAgent: string, language: string, cores: number): Promise<string> {
  const data = `${userAgent}-${language}-${cores}`;
  const encoder = new TextEncoder();
  const dataBuffer = encoder.encode(data);
  
  const hashBuffer = await crypto.subtle.digest('SHA-256', dataBuffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
  
  return hashHex;
}

export class SessionGuard {
  constructor(private expectedFingerprint: string) {}

  async validateCurrentSession(userAgent: string, language: string, cores: number): Promise<boolean> {
    const currentFingerprint = await generateHardwareFingerprint(userAgent, language, cores);
    return currentFingerprint === this.expectedFingerprint;
  }
}
