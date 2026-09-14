import { describe, it, expect } from 'vitest';
import { generateHardwareFingerprint, SessionGuard } from '../src/pro/session-guard';

describe('SessionGuard Cryptographic Fingerprint', () => {
  it('should generate a consistent 64-character hex hash', async () => {
    const hash = await generateHardwareFingerprint('Mozilla/5.0 Node', 'en-US', 8);
    expect(hash).toHaveLength(64); // SHA-256 is 32 bytes = 64 hex characters
    expect(hash).toMatch(/^[0-9a-f]{64}$/);
    
    const hash2 = await generateHardwareFingerprint('Mozilla/5.0 Node', 'en-US', 8);
    expect(hash).toBe(hash2); // Should be deterministic for the same hardware
  });

  it('should validate matching sessions', async () => {
    const hash = await generateHardwareFingerprint('Chrome/100', 'es-ES', 4);
    const guard = new SessionGuard(hash);
    
    const isValid = await guard.validateCurrentSession('Chrome/100', 'es-ES', 4);
    expect(isValid).toBe(true);
  });

  it('should reject mismatched sessions', async () => {
    const hash = await generateHardwareFingerprint('Chrome/100', 'es-ES', 4);
    const guard = new SessionGuard(hash);
    
    const isValid = await guard.validateCurrentSession('Chrome/100', 'en-US', 4);
    expect(isValid).toBe(false);
  });
});
