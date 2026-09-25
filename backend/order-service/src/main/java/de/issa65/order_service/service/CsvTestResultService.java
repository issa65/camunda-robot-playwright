package de.issa65.order_service.service;

import org.springframework.stereotype.Service;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.Instant;

@Service
public class CsvTestResultService {

    private static final Path CSV_PATH =
            Path.of("data", "test-results.csv");

    private static final String HEADER =
            "timestamp;" +
                    "processInstanceKey;" +
                    "testRunId;" +
                    "testName;" +
                    "testType;" +
                    "component;" +
                    "workerUser;" +
                    "workerHost;" +
                    "workerIpAddress;" +
                    "result;" +
                    "status;" +
                    "durationSeconds;" +
                    "deliveryDurationSeconds;" +
                    "startupToLoadingSeconds;" +
                    "loadingToMainSeconds;" +
                    "startupToMainSeconds;" +
                    "startupPath;" +
                    "errorType;" +
                    "errorMessage;" +
                    "message";

    public synchronized void append(
            long processInstanceKey,
            String testRunId,
            String testName,
            String testType,
            String component,

            String workerUser,
            String workerHost,
            String workerIpAddress,

            String result,
            String status,

            Double durationSeconds,
            Double deliveryDurationSeconds,

            Double startupToLoadingSeconds,
            Double loadingToMainSeconds,
            Double startupToMainSeconds,
            String startupPath,

            String errorType,
            String errorMessage,
            String message
    ) throws IOException {

        Files.createDirectories(CSV_PATH.getParent());

        if (Files.notExists(CSV_PATH)) {
            Files.writeString(
                    CSV_PATH,
                    HEADER + System.lineSeparator(),
                    StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE
            );
        }

        String timestamp = Instant.now().toString();

        String line = String.join(";",
                csv(timestamp),
                String.valueOf(processInstanceKey),

                csv(testRunId),
                csv(testName),
                csv(testType),
                csv(component),

                csv(workerUser),
                csv(workerHost),
                csv(workerIpAddress),

                csv(result),
                csv(status),

                number(durationSeconds),
                number(deliveryDurationSeconds),

                number(startupToLoadingSeconds),
                number(loadingToMainSeconds),
                number(startupToMainSeconds),

                csv(startupPath),

                csv(errorType),
                csv(errorMessage),
                csv(message)
        );

        Files.writeString(
                CSV_PATH,
                line + System.lineSeparator(),
                StandardCharsets.UTF_8,
                StandardOpenOption.CREATE,
                StandardOpenOption.APPEND
        );
    }

    private String number(Double value) {
        if (value == null) {
            return "";
        }

        return String.valueOf(value);
    }

    private String csv(String value) {
        if (value == null) {
            return "";
        }

        String escaped = value.replace("\"", "\"\"");

        if (escaped.contains(";")
                || escaped.contains("\"")
                || escaped.contains("\n")
                || escaped.contains("\r")) {

            return "\"" + escaped + "\"";
        }

        return escaped;
    }
}