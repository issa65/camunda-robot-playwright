import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

import createCamundaClient
  from "@camunda8/orchestration-cluster-api";


// ============================================================
// Paths
// ============================================================

const __filename = fileURLToPath(
  import.meta.url
);

const __dirname = path.dirname(
  __filename
);

// worker.js liegt unter:
// playwright-worker/src/worker.js
//
// Projektverzeichnis ist also eine Ebene höher.
const workerDirectory = path.resolve(
  __dirname,
  ".."
);


// ============================================================
// Camunda connection
// ============================================================

const camunda = createCamundaClient({
  config: {
    CAMUNDA_REST_ADDRESS:
      process.env.CAMUNDA_REST_ADDRESS
      ?? "http://localhost:8080",

    CAMUNDA_AUTH_STRATEGY:
      "NONE"
  }
});


console.log(
  "Starting Camunda Playwright worker..."
);


// ============================================================
// Connection check
// ============================================================

try {

  const topology =
    await camunda.getTopology();

  console.log(
    "Connected to Camunda."
  );

  console.log(
    `Brokers: ${topology.brokers?.length ?? 0}`
  );

} catch (error) {

  console.error(
    "Could not connect to Camunda:",
    error
  );

  process.exit(1);
}


// ============================================================
// Playwright test runner
// ============================================================

function runPlaywrightLoginTest() {

  return new Promise(
    (resolve, reject) => {

      console.log(
        "Starting Playwright test..."
      );

      const command =
        process.env.ComSpec || "cmd.exe";


      const testProcess = spawn(
        command,
        [
          "/d",
          "/s",
          "/c",
          "npx playwright test tests/login.spec.cjs --headed --reporter=line"
        ],
        {
          cwd: workerDirectory,
          stdio: "inherit"
        }
      );


      testProcess.on(
        "close",
        (exitCode) => {

          if (exitCode === 0) {

            console.log(
              "Playwright test completed successfully."
            );

            resolve();

            return;
          }


          reject(
            new Error(
              "Playwright test failed with "
              + `exit code ${exitCode}.`
            )
          );
        }
      );


      testProcess.on(
        "error",
        (error) => {

          reject(
            new Error(
              "Could not start Playwright test: "
              + error.message
            )
          );
        }
      );
    }
  );
}


// ============================================================
// Camunda Job Worker
// ============================================================

camunda.createJobWorker({

  jobType:
    "playwright-login-check",

  workerName:
    "playwright-worker",

  maxParallelJobs:
    1,

  jobTimeoutMs:
    60_000,

  jobHandler: async (job) => {

    console.log(
      `Received job: ${job.jobKey}`
    );

    console.log(
      "Variables:",
      job.variables
    );


    const startTime =
      performance.now();


    try {

      // ------------------------------------------------------
      // Run existing Playwright test
      // ------------------------------------------------------

      await runPlaywrightLoginTest();


      // ------------------------------------------------------
      // Duration
      // ------------------------------------------------------

      const durationSeconds =
        Math.round(
          (
            performance.now()
            - startTime
          ) / 10
        ) / 100;


      console.log(
        "Playwright Service Task PASS"
      );

      console.log(
        `Duration: ${durationSeconds}s`
      );


      // ------------------------------------------------------
      // Complete Camunda job
      // ------------------------------------------------------

      return job.complete({

        testResult:
          "PASS",

        status:
          "FUNCTIONAL_OK",

        durationSeconds:
          durationSeconds,

        playwrightResult:
          "PASS"

      });


    } catch (error) {

      const durationSeconds =
        Math.round(
          (
            performance.now()
            - startTime
          ) / 10
        ) / 100;


      console.error(
        "Playwright Service Task FAIL"
      );

      console.error(
        error
      );

      console.error(
        `Duration: ${durationSeconds}s`
      );


      // Kein falsches PASS an Camunda senden.
      // Fehler wird an den Job Worker weitergegeben.
      throw error;
    }
  }
});


console.log(
  "Waiting for jobs of type: "
  + "playwright-login-check"
);