package de.issa65.order_service.worker;

import de.issa65.order_service.service.CsvTestResultService;
import io.camunda.client.annotation.JobWorker;
import io.camunda.client.api.response.ActivatedJob;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.util.List;
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
                    "result",
                    "status",
                    "durationSeconds",
                    "message",
                    "metrics"
            }
    )
    public void saveTestResult(ActivatedJob job) throws IOException {

        Map<String, Object> variables = job.getVariablesAsMap();

        String testRunId = (String) variables.get("testRunId");
        String testName = (String) variables.get("testName");
        String testType = (String) variables.get("testType");
        String component = (String) variables.get("component");
        String result = (String) variables.get("result");
        String status = (String) variables.get("status");
        String message = (String) variables.get("message");

        Double durationSeconds =
                toDouble(variables.get("durationSeconds"));

        @SuppressWarnings("unchecked")
        List<Map<String, Object>> metrics =
                (List<Map<String, Object>>) variables.get("metrics");

        csvTestResultService.append(
                job.getProcessInstanceKey(),
                testRunId,
                testName,
                testType,
                component,
                result,
                status,
                durationSeconds,
                message,
                metrics
        );
    }

    private Double toDouble(Object value) {
        if (value == null) {
            return null;
        }

        if (value instanceof Number number) {
            return number.doubleValue();
        }

        return Double.valueOf(value.toString());
    }
}