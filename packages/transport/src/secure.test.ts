import { describe, expect, it } from "vitest";
import { createPairIdentity, createSecureChannel, pairConnectionUrl } from "./secure.js";

async function connected(secret = createPairIdentity().secret, other = secret) {
  const leftWire: string[] = [];
  const rightWire: string[] = [];
  const received: string[] = [];
  const left = createSecureChannel({
    secret,
    role: "addin",
    context: "test",
    send: (s) => leftWire.push(s),
    receive: (s) => received.push(s),
  });
  const right = createSecureChannel({
    secret: other,
    role: "agent",
    context: "test",
    send: (s) => rightWire.push(s),
    receive: (s) => received.push(s),
  });
  await left.start();
  await right.start();
  const pump = async () => {
    for (let i = 0; i < 8 && (leftWire.length || rightWire.length); i++) {
      for (const frame of leftWire.splice(0)) await right.accept(frame);
      for (const frame of rightWire.splice(0)) await left.accept(frame);
    }
  };
  await pump();
  return { left, right, leftWire, rightWire, received, pump };
}

describe("end-to-end channel", () => {
  it("authenticates both peers and encrypts workbook content in both directions", async () => {
    const { left, right, leftWire, received, pump } = await connected();
    expect(left.authenticated).toBe(true);
    expect(right.authenticated).toBe(true);
    await left.send("cells: salary 12345");
    expect(leftWire.join()).not.toContain("salary");
    await pump();
    await right.send("reply: confidential");
    await pump();
    expect(received).toEqual(["cells: salary 12345", "reply: confidential"]);
  });

  it("rejects a peer that knows the UID but not the secret", async () => {
    await expect(
      connected(createPairIdentity().secret, createPairIdentity().secret),
    ).rejects.toThrow();
  });

  it("rejects tampering, reflection, plaintext, and replay", async () => {
    const { left, right, leftWire, received } = await connected();
    await left.send("private cells");
    const frame = leftWire.shift() ?? "";
    await expect(left.accept(frame)).rejects.toThrow();
    await expect(right.accept('{"type":"request","method":"excel.range.read"}')).rejects.toThrow();
    const altered = JSON.parse(frame);
    altered.ciphertext = `${altered.ciphertext[0] === "A" ? "B" : "A"}${altered.ciphertext.slice(1)}`;
    await expect(right.accept(JSON.stringify(altered))).rejects.toThrow();
    await right.accept(frame);
    await expect(right.accept(frame)).rejects.toThrow();
    expect(received).toEqual(["private cells"]);
  });

  it("rejects captured traffic in a fresh connection", async () => {
    const secret = createPairIdentity().secret;
    const old = await connected(secret);
    await old.left.send("captured");
    const fresh = await connected(secret);
    await expect(fresh.right.accept(old.leftWire[0] ?? "")).rejects.toThrow();
    expect(fresh.received).toEqual([]);
  });

  it("never puts the encryption secret in relay query parameters", async () => {
    const identity = createPairIdentity();
    const url = new URL(await pairConnectionUrl("https://pair.example.test", identity, "agent"));
    expect(url.pathname).toBe("/connect");
    expect(url.search).not.toContain(identity.secret);
    expect(url.hash).toBe(`#${identity.secret}`);
    expect(url.searchParams.get("uid")).toBe(identity.id);
  });
});

it("encrypts a bounded large result without overflowing the encoding stack", async () => {
  const { left, received, pump } = await connected();
  const content = "synthetic ".repeat(50_000);
  await left.send(content);
  await pump();
  expect(received).toEqual([content]);
  await expect(left.send("x".repeat(1_000_001))).rejects.toThrow("limit");
});
