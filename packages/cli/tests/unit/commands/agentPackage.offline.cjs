// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Fault seams only: the actual CLI, public core, schemas and native template remain real.
let violated = false;
function forbidden() {
  violated = true;
  process.stderr.write("LOCAL_AGENT_OFFLINE_BOUNDARY_VIOLATION\n");
  throw new Error("A local agent command attempted network, process execution, or a CWD change.");
}

require("net").Socket.prototype.connect = forbidden;
require("http").request = forbidden;
require("http").get = forbidden;
require("https").request = forbidden;
require("https").get = forbidden;
global.fetch = forbidden;
process.chdir = forbidden;
for (const method of ["spawn", "spawnSync", "exec", "execSync", "execFile", "execFileSync"]) {
  require("child_process")[method] = forbidden;
}
process.on("exit", () => {
  if (violated) process.exitCode = 97;
});

// Emitting after listener registration exercises cooperative cancellation on Windows too,
// where child.kill("SIGINT") terminates the process instead of delivering console Ctrl+C.
if (process.env.ATK_TEST_CANCEL_LOCAL_AGENT === "true") {
  const on = process.on;
  process.on = function (event, listener) {
    const result = on.call(this, event, listener);
    if (event === "SIGINT") queueMicrotask(() => process.emit("SIGINT"));
    return result;
  };
}
