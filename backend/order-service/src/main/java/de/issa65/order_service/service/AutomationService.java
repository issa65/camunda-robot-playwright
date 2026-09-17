package de.issa65.order_service.service;

import io.camunda.client.CamundaClient;
import io.camunda.client.api.response.ProcessInstanceEvent;
import org.springframework.stereotype.Service;

import java.util.Map;

@Service
public class AutomationService {

    private final CamundaClient camundaClient;

    public AutomationService(CamundaClient camundaClient) {
        this.camundaClient = camundaClient;
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
}