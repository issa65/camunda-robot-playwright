package de.issa65.order_service.api;

import jakarta.validation.constraints.NotBlank;

public record AutomationRequest(
        @NotBlank String url
) {
}