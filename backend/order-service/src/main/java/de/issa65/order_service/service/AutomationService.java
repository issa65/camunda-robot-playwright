package de.issa65.order_service.service;

import de.issa65.order_service.api.AutomationStatusResponse;
import io.camunda.client.CamundaClient;
import io.camunda.client.api.response.ProcessInstanceEvent;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestClient;
import tools.jackson.databind.ObjectMapper;
import org.springframework.web.client.RestClientResponseException;

import java.util.List;
import java.util.Map;

@Service
public class AutomationService {

    private final CamundaClient camundaClient;
    private final RestClient restClient;
    private final ObjectMapper objectMapper;

    public AutomationService(
            CamundaClient camundaClient,
            RestClient.Builder restClientBuilder,
            ObjectMapper objectMapper) {

        this.camundaClient = camundaClient;
        this.objectMapper = objectMapper;

        this.restClient = restClientBuilder
                .baseUrl("http://localhost:8080/v2")
                .build();
    }

    public ProcessInstanceEvent startAutomation(String url) {
        return camundaClient
                .newCreateInstanceCommand()
                .bpmnProcessId("rpa-smoke-test")
                .latestVersion()
                .variables(Map.of("url", url))
                .send()
                .join();
    }

    public AutomationStatusResponse getStatus(String processInstanceKey) {

        Map<?, ?> response;

        try {
            response = restClient
                    .get()
                    .uri("/process-instances/{key}", processInstanceKey)
                    .retrieve()
                    .body(Map.class);

        } catch (RestClientResponseException exception) {

            if (exception.getStatusCode().value() == 404) {
                throw new AutomationNotFoundException(processInstanceKey);
            }

            throw exception;
        }

        if (response == null) {
            throw new IllegalStateException("No response from Camunda");
        }

        String pageTitle = getPageTitle(processInstanceKey);

        return new AutomationStatusResponse(
                String.valueOf(response.get("processInstanceKey")),
                String.valueOf(response.get("processDefinitionId")),
                String.valueOf(response.get("state")),
                Boolean.TRUE.equals(response.get("hasIncident")),
                pageTitle
        );
    }

    private String getPageTitle(String processInstanceKey) {

        Map<?, ?> response = restClient
                .post()
                .uri("/variables/search")
                .body(Map.of(
                        "filter", Map.of(
                                "processInstanceKey", processInstanceKey,
                                "name", "pageTitle"
                        ),
                        "page", Map.of(
                                "limit", 1
                        )
                ))
                .retrieve()
                .body(Map.class);

        if (response == null) {
            return null;
        }

        Object itemsObject = response.get("items");

        if (!(itemsObject instanceof List<?> items) || items.isEmpty()) {
            return null;
        }

        Object firstItem = items.get(0);

        if (!(firstItem instanceof Map<?, ?> variable)) {
            return null;
        }

        Object value = variable.get("value");

        if (value == null) {
            return null;
        }

        String rawValue = String.valueOf(value);

        try {
            return objectMapper.readValue(rawValue, String.class);
        } catch (Exception e) {
            return rawValue;
        }
    }

    public String completeUserTask(String processInstanceKey) {

        Map<?, ?> response = restClient
                .post()
                .uri("/user-tasks/search")
                .body(Map.of(
                        "filter", Map.of(
                                "processInstanceKey", processInstanceKey,
                                "state", "CREATED"
                        ),
                        "page", Map.of(
                                "limit", 10
                        )
                ))
                .retrieve()
                .body(Map.class);

        if (response == null) {
            throw new IllegalStateException("No response from Camunda");
        }

        Object itemsObject = response.get("items");

        if (!(itemsObject instanceof List<?> items) || items.isEmpty()) {
            throw new UserTaskNotFoundException(processInstanceKey);
        }

        Object firstItem = items.get(0);

        if (!(firstItem instanceof Map<?, ?> task)) {
            throw new IllegalStateException("Invalid user task response");
        }

        String userTaskKey = String.valueOf(task.get("userTaskKey"));

        restClient
                .post()
                .uri("/user-tasks/{userTaskKey}/completion", userTaskKey)
                .body(Map.of())
                .retrieve()
                .toBodilessEntity();

        return userTaskKey;
    }


}