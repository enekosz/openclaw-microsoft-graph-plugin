import { afterEach, describe, expect, it, vi } from "vitest";
import { clearMicrosoftAccountIdentityCacheForTests, verifyMicrosoftAccountIdentity } from "./index.js";

afterEach(() => {
  clearMicrosoftAccountIdentityCacheForTests();
  vi.restoreAllMocks();
});

describe("Microsoft delegated account binding", () => {
  it("accepts the exact user principal name and caches the verified token identity", async () => {
    const reader = vi.fn(async () => ({
      id: "immutable-user-id",
      userPrincipalName: "Eneko@Acupuntura.Live",
      mail: "eneko@acupuntura.live",
    }));
    const signal = new AbortController().signal;

    const first = await verifyMicrosoftAccountIdentity("eneko@acupuntura.live", "token-one", signal, reader);
    const second = await verifyMicrosoftAccountIdentity("ENEKO@ACUPUNTURA.LIVE", "token-one", signal, reader);

    expect(first).toEqual({
      id: "immutable-user-id",
      userPrincipalName: "eneko@acupuntura.live",
      mail: "eneko@acupuntura.live",
      verifiedAs: "eneko@acupuntura.live",
    });
    expect(second).toEqual(first);
    expect(reader).toHaveBeenCalledTimes(1);
    expect(reader).toHaveBeenCalledWith("token-one", "/me?$select=id,userPrincipalName,mail", { signal });
  });

  it("accepts an exact mail alias when the tenant UPN differs", async () => {
    const identity = await verifyMicrosoftAccountIdentity(
      "eneko@acupuntura.live",
      "token-alias",
      new AbortController().signal,
      async () => ({ id: "user-id", userPrincipalName: "eneko@tenant.onmicrosoft.com", mail: "eneko@acupuntura.live" }),
    );
    expect(identity.verifiedAs).toBe("eneko@acupuntura.live");
    expect(identity.userPrincipalName).toBe("eneko@tenant.onmicrosoft.com");
  });

  it("fails closed on the wrong account and does not cache the refusal", async () => {
    const reader = vi.fn(async () => ({ id: "wrong-user", userPrincipalName: "admin@theasx.company", mail: "admin@theasx.company" }));
    const signal = new AbortController().signal;

    await expect(verifyMicrosoftAccountIdentity("eneko@acupuntura.live", "wrong-token", signal, reader)).rejects.toThrow("credential_account_mismatch");
    await expect(verifyMicrosoftAccountIdentity("eneko@acupuntura.live", "wrong-token", signal, reader)).rejects.toThrow("credential_account_mismatch");
    expect(reader).toHaveBeenCalledTimes(2);
  });

  it("rejects malformed expected identities and incomplete provider profiles", async () => {
    const signal = new AbortController().signal;
    await expect(verifyMicrosoftAccountIdentity("not-an-account", "token", signal, vi.fn())).rejects.toThrow("invalid_expected_user_principal_name");
    await expect(verifyMicrosoftAccountIdentity("eneko@acupuntura.live", "token", signal, async () => ({ id: "user-id" }))).rejects.toThrow("invalid_provider_response");
  });
});
