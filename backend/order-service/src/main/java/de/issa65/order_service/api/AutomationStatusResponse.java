package de.issa65.order_service.api;

public record AutomationStatusResponse(
        String processInstanceKey,
        String processDefinitionId,
        String state,
        boolean hasIncident,
        String pageTitle
) {
}