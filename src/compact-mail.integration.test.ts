import { describe, expect, it, vi } from "vitest";

vi.mock("./credential.js", () => ({
  readCredential: vi.fn(async () => ({ clientId: "synthetic", refreshToken: "synthetic", tenant: "common", scopes: ["Mail.Read", "User.Read"] })),
  selectScope: (_credential: unknown, allowed: string[]) => allowed[0],
  exchangeRefreshToken: vi.fn(async () => "synthetic-compact-mail-token"),
  tokenForAuthorizedOperation: vi.fn(async () => "synthetic-compact-mail-token"),
}));

import { clearMicrosoftAccountIdentityCacheForTests, executeCompactMicrosoftRead } from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

const params = { action: "list_messages", folder: "inbox", isRead: false, orderBy: "receivedDateTime", orderDirection: "desc", includeBody: false, limit: 10 };
const request = { toolCallId: "synthetic-mail-call", toolName: "outlook_mail_read" as const, agentId: "main", params };

describe("bounded compact inbox mail read", () => {
  it("uses only GET, binds the account, preserves the query and reports continuation", async () => {
    clearMicrosoftAccountIdentityCacheForTests();
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async (input, init) => {
      expect(init?.method ?? "GET").toBe("GET");
      const url = new URL(String(input));
      if (url.pathname === "/v1.0/me") return new Response(JSON.stringify({ id: "synthetic-owner", mail: "synthetic@example.test", userPrincipalName: "synthetic@example.test" }), { status: 200 });
      expect(url.pathname).toBe("/v1.0/me/mailFolders/inbox/messages");
      expect(url.searchParams.get("$top")).toBe("10");
      expect(url.searchParams.get("$orderby")).toBe("receivedDateTime desc");
      expect(url.searchParams.get("$filter")).toContain("isRead eq false");
      expect(url.searchParams.get("$select")?.split(",")).not.toContain("body");
      return new Response(JSON.stringify({ value: [{ id: "synthetic-message", subject: "Synthetic message", isRead: false }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skiptoken=synthetic" }), { status: 200 });
    });
    try {
      const result = await executeCompactMicrosoftRead({ enabled: true, expectedUserPrincipalName: "synthetic@example.test", policy: graphPolicyFixture() } as any, request) as any;
      expect(result).toMatchObject({ ok: true, action: "list_messages", mailboxQuery: params, microsoftAccount: { verifiedAs: "synthetic@example.test" }, items: [{ subject: "Synthetic message", isRead: false }], truncated: true });
      expect(fetchSpy).toHaveBeenCalledTimes(2);
      expect(result).not.toHaveProperty("mutationApplied", true);
    } finally { fetchSpy.mockRestore(); }
  });

  it("rejects mutation, expanded reads and malformed criteria before network access", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    try {
      for (const changes of [{ action: "mark_read" }, { folder: "sentitems" }, { isRead: true }, { includeBody: true }, { orderDirection: "asc" }, { limit: 26 }, { limit: 0 }, { limit: "10" }, { messageId: "synthetic" }, { continuation: "synthetic" }]) {
        await expect(executeCompactMicrosoftRead({} as any, { ...request, params: { ...params, ...changes } })).rejects.toThrow("invalid_compact_read_request");
      }
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally { fetchSpy.mockRestore(); }
  });

  it("retains default-deny service policy before credential or provider access", async () => {
    const policy = graphPolicyFixture();
    policy.services!.mail!.agents!.main!.operations = [];
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    try {
      await expect(executeCompactMicrosoftRead({ enabled: true, policy } as any, request)).rejects.toThrow();
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally { fetchSpy.mockRestore(); }
  });
});
