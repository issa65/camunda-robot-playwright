package de.issa65.order_service.worker;

import de.issa65.order_service.service.CsvTestResultService;
import io.camunda.client.annotation.JobWorker;
import io.camunda.client.api.response.ActivatedJob;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.util.Map;

@Component
public class TestResultWorker {

    private final CsvTestResultService csvTestResultService;

    public TestResultWorker(CsvTestResultService csvTestResultService) {
        this.csvTestResultService = csvTestResultService;
    }

    @JobWorker(
            type = "save-test-result",
            fetchVariables = {
                    "testRunId",
                    "testName",
                    "testType",
                    "component",

                    "workerUser",
                    "workerHost",
                    "workerIpAddress",

                    "result",
                    "status",

                    "durationSeconds",
                    "deliveryDurationSeconds",

                    "startupToLoadingSeconds",
                    "loadingToMainSeconds",
                    "startupToMainSeconds",
                    "startupPath",

                    "errorType",
                    "errorMessage",
                    "message"
            }
    )
    public void saveTestResult(ActivatedJob job) throws IOException {

        Map<String, Object> variables =
                job.getVariablesAsMap();

        String testRunId =
                toStringValue(variables.get("testRunId"));

        String testName =
                toStringValue(variables.get("testName"));

        String testType =
                toStringValue(variables.get("testType"));

        String component =
                toStringValue(variables.get("component"));

        String workerUser =
                toStringValue(variables.get("workerUser"));

        String workerHost =
                toStringValue(variables.get("workerHost"));

        String workerIpAddress =
                toStringValue(variables.get("workerIpAddress"));

        String result =
                toStringValue(variables.get("result"));

        String status =
                toStringValue(variables.get("status"));

        Double durationSeconds =
                toDouble(variables.get("durationSeconds"));

        Double deliveryDurationSeconds =
                toDouble(variables.get("deliveryDurationSeconds"));

        Double startupToLoadingSeconds =
                toDouble(variables.get("startupToLoadingSeconds"));

        Double loadingToMainSeconds =
                toDouble(variables.get("loadingToMainSeconds"));

        Double startupToMainSeconds =
                toDouble(variables.get("startupToMainSeconds"));

        String startupPath =
                toStringValue(variables.get("startupPath"));

        String errorType =
                toStringValue(variables.get("errorType"));

        String errorMessage =
                toStringValue(variables.get("errorMessage"));

        String message =
                toStringValue(variables.get("message"));

        csvTestResultService.append(
                job.getProcessInstanceKey(),

                testRunId,
                testName,
                testType,
                component,

                workerUser,
                workerHost,
                workerIpAddress,

                result,
                status,

                durationSeconds,
                deliveryDurationSeconds,

                startupToLoadingSeconds,
                loadingToMainSeconds,
                startupToMainSeconds,
                startupPath,

                errorType,
                errorMessage,
                message
        );
    }

    private Double toDouble(Object value) {
        if (value == null) {
            return null;
        }

        if (value instanceof Number number) {
            return number.doubleValue();
        }

        String text = value.toString();

        if (text.isBlank()) {
            return null;
        }

        return Double.valueOf(text);
    }

    private String toStringValue(Object value) {
        if (value == null) {
            return "";
        }

        return value.toString();
    }
}