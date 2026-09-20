package de.issa65.order_service.service;

import org.springframework.stereotype.Service;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.Instant;
import java.util.List;
import java.util.Map;

@Service
public class CsvTestResultService {

    private static final Path CSV_PATH =
            Path.of("data", "test-results.csv");

    private static final String HEADER =
            "timestamp;processInstanceKey;testRunId;testName;testType;component;" +
                    "recordType;result;status;durationSeconds;metricName;metricValue;" +
                    "metricUnit;message";

    public synchronized void append(
            long processInstanceKey,
            String testRunId,
            String testName,
            String testType,
            String component,
            String result,
            String status,
            Double durationSeconds,
            String message,
            List<Map<String, Object>> metrics
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

        writeLine(
                timestamp,
                processInstanceKey,
                testRunId,
                testName,
                testType,
                component,
                "SUMMARY",
                result,
                status,
                durationSeconds,
                "",
                "",
                "",
                message
        );

        if (metrics != null) {
            for (Map<String, Object> metric : metrics) {
                writeLine(
                        timestamp,
                        processInstanceKey,
                        testRunId,
                        testName,
                        testType,
                        component,
                        "METRIC",
                        "",
                        "",
                        null,
                        String.valueOf(metric.getOrDefault("name", "")),
                        String.valueOf(metric.getOrDefault("value", "")),
                        String.valueOf(metric.getOrDefault("unit", "")),
                        ""
                );
            }
        }
    }

    private void writeLine(
            String timestamp,
            long processInstanceKey,
            String testRunId,
            String testName,
            String testType,
            String component,
            String recordType,
            String result,
            String status,
            Double durationSeconds,
            String metricName,
            String metricValue,
            String metricUnit,
            String message
    ) throws IOException {

        String line = String.join(";",
                csv(timestamp),
                String.valueOf(processInstanceKey),
                csv(testRunId),
                csv(testName),
                csv(testType),
                csv(component),
                csv(recordType),
                csv(result),
                csv(status),
                durationSeconds == null ? "" : String.valueOf(durationSeconds),
                csv(metricName),
                csv(metricValue),
                csv(metricUnit),
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