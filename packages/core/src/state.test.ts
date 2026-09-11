import { describe, expect, it } from "vitest";
import { SessionStateMachine } from "./state.js";

describe("SessionStateMachine", () => {
  it("accepts the read execution lifecycle", () => {
    const machine = new SessionStateMachine();
    machine.transition("running_agent");
    machine.transition("executing_read_tool");
    machine.transition("running_agent");
    machine.transition("completed");
    expect(machine.state).toBe("completed");
  });

  it("rejects invalid transitions", () => {
    const machine = new SessionStateMachine();
    expect(() => machine.transition("executing_write_tool")).toThrow(/Cannot transition/);
  });
});
