package de.issa65.order_service.api;

import java.time.Instant;

public record ApiErrorResponse(
        Instant timestamp,
        int status,
        String error,
        String message,
        String path
) {
}