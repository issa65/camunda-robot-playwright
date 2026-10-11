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

// worker.js:
// playwright-worker/src/worker.js
//
// Root:
// playwright-worker/
const workerDirectory = path.resolve(
  __dirname,
  ".."
);


// ============================================================
// Playwright test registry
// ============================================================
//
// Camunda kennt nur den fachlichen testName.
// Der Worker übersetzt diesen in die konkrete Spec-Datei.
//
// Neue Tests werden später nur hier ergänzt.
//

const playwrightTests = {

  login:
    "tests/login.spec.cjs",

  orderEntry:
    "tests/order-entry.spec.cjs"

};


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
// Camunda connection check
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
// Generic Playwright test runner
// ============================================================

function runPlaywrightTest(testFile) {

  return new Promise(
    (resolve, reject) => {

      const command =
        process.env.ComSpec
        || "cmd.exe";


      console.log(
        `Starting Playwright test: ${testFile}`
      );


      const testProcess = spawn(
        command,
        [
          "/d",
          "/s",
          "/c",
          `npx playwright test ${testFile} --headed --reporter=line`
        ],
        {
          cwd:
            workerDirectory,

          stdio:
            "inherit"
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
// Generic Camunda Playwright Worker
// ============================================================
//
// Alle Playwright Tests verwenden denselben Job Type:
//
// playwright-test
//
// Welcher Test ausgeführt wird, entscheidet:
//
// job.variables.testName
//

camunda.createJobWorker({

  jobType:
    "playwright-test",

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


    // --------------------------------------------------------
    // Read test name from Camunda
    // --------------------------------------------------------

    const testName =
      job.variables?.testName;


    if (!testName) {

      throw new Error(
        "Missing required Camunda variable: testName"
      );
    }


    // --------------------------------------------------------
    // Resolve Playwright test
    // --------------------------------------------------------

    const testFile =
      playwrightTests[testName];


    if (!testFile) {

      throw new Error(
        `Unknown Playwright test: ${testName}. `
        + "Available tests: "
        + Object.keys(playwrightTests).join(", ")
      );
    }


    console.log(
      `Test name: ${testName}`
    );

    console.log(
      `Test file: ${testFile}`
    );


    // --------------------------------------------------------
    // Start duration measurement
    // --------------------------------------------------------

    const startTime =
      performance.now();


    try {

      // ------------------------------------------------------
      // Execute Playwright test
      // ------------------------------------------------------

      await runPlaywrightTest(
        testFile
      );


      // ------------------------------------------------------
      // Calculate duration
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
        `Test: ${testName}`
      );

      console.log(
        `Duration: ${durationSeconds}s`
      );


      // ------------------------------------------------------
      // Complete Camunda Job
      // ------------------------------------------------------

      return job.complete({

        testName:
          testName,

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

      // ------------------------------------------------------
      // Calculate duration on failure
      // ------------------------------------------------------

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
        `Test: ${testName}`
      );

      console.error(
        `Duration: ${durationSeconds}s`
      );

      console.error(
        error
      );


      // Fehler an Camunda weitergeben.
      // Kein falsches PASS erzeugen.
      throw error;
    }
  }
});


// ============================================================
// Worker ready
// ============================================================

console.log(
  "Waiting for jobs of type: playwright-test"
);

console.log(
  "Available Playwright tests:",
  Object.keys(playwrightTests)
);