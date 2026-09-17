package de.issa65.order_service.api;

import de.issa65.order_service.service.AutomationService;
import io.camunda.client.api.response.ProcessInstanceEvent;
import jakarta.validation.Valid;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import de.issa65.order_service.api.AutomationStatusResponse;

import java.util.Map;

@RestController
@RequestMapping("/api/automation")
public class AutomationController {

    private final AutomationService automationService;

    public AutomationController(AutomationService automationService) {
        this.automationService = automationService;
    }

    @PostMapping("/start")
    public ResponseEntity<Map<String, Object>> start(
            @Valid @RequestBody AutomationRequest request) {

        ProcessInstanceEvent instance =
                automationService.startAutomation(request.url());

        return ResponseEntity.ok(Map.of(
                "processInstanceKey", instance.getProcessInstanceKey(),
                "bpmnProcessId", instance.getBpmnProcessId(),
                "version", instance.getVersion()
        ));
    }

    @GetMapping("/{processInstanceKey}")
    public ResponseEntity<AutomationStatusResponse> getStatus(
            @PathVariable String processInstanceKey) {

        return ResponseEntity.ok(
                automationService.getStatus(processInstanceKey)
        );
    }
}