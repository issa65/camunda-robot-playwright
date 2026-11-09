
package de.issa65.order_service.api;

import de.issa65.order_service.service.AutomationNotFoundException;
import de.issa65.order_service.service.UserTaskNotFoundException;

import jakarta.servlet.http.HttpServletRequest;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.MethodArgumentNotValidException;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.client.RestClientException;

import java.time.Instant;

@RestControllerAdvice
public class GlobalExceptionHandler {

    private static final Logger log =
            LoggerFactory.getLogger(GlobalExceptionHandler.class);

    // ============================================================
    // HTTP 400 - Validation Error
    // ============================================================

    @ExceptionHandler(MethodArgumentNotValidException.class)
    public ResponseEntity<ApiErrorResponse> handleValidationError(
            MethodArgumentNotValidException exception,
            HttpServletRequest request) {

        String message = exception.getBindingResult()
                .getFieldErrors()
                .stream()
                .findFirst()
                .map(error ->
                        error.getField() + " " + error.getDefaultMessage()
                )
                .orElse("Invalid request");

        log.warn(
                "Validation error on {}: {}",
                request.getRequestURI(),
                message
        );

        return buildResponse(
                HttpStatus.BAD_REQUEST,
                message,
                request.getRequestURI()
        );
    }

    // ============================================================
    // HTTP 404 - Automation Not Found
    // ============================================================

    @ExceptionHandler(AutomationNotFoundException.class)
    public ResponseEntity<ApiErrorResponse> handleAutomationNotFound(
            AutomationNotFoundException exception,
            HttpServletRequest request) {

        log.warn(
                "Automation not found on {}: {}",
                request.getRequestURI(),
                exception.getMessage()
        );

        return buildResponse(
                HttpStatus.NOT_FOUND,
                exception.getMessage(),
                request.getRequestURI()
        );
    }

    // ============================================================
    // HTTP 409 - User Task Conflict
    // ============================================================

    @ExceptionHandler(UserTaskNotFoundException.class)
    public ResponseEntity<ApiErrorResponse> handleUserTaskNotFound(
            UserTaskNotFoundException exception,
            HttpServletRequest request) {

        log.warn(
                "User task conflict on {}: {}",
                request.getRequestURI(),
                exception.getMessage()
        );

        return buildResponse(
                HttpStatus.CONFLICT,
                exception.getMessage(),
                request.getRequestURI()
        );
    }

    // ============================================================
    // HTTP 502 - Camunda REST Communication Error
    // ============================================================

    @ExceptionHandler(RestClientException.class)
    public ResponseEntity<ApiErrorResponse> handleCamundaError(
            RestClientException exception,
            HttpServletRequest request) {

        log.error(
                "REST communication error on {}",
                request.getRequestURI(),
                exception
        );

        return buildResponse(
                HttpStatus.BAD_GATEWAY,
                "Communication with Camunda failed",
                request.getRequestURI()
        );
    }


    // ============================================================
    // HTTP 409 - Duplicate Order
    // ============================================================

    @ExceptionHandler(IllegalArgumentException.class)
    public ResponseEntity<ApiErrorResponse> handleDuplicateOrder(
            IllegalArgumentException exception,
            HttpServletRequest request) {

        log.warn(
                "Order conflict on {}: {}",
                request.getRequestURI(),
                exception.getMessage()
        );

        return buildResponse(
                HttpStatus.CONFLICT,
                exception.getMessage(),
                request.getRequestURI()
        );
    }


    // ============================================================
    // HTTP 500 - Unexpected Error
    // ============================================================

    @ExceptionHandler(Exception.class)
    public ResponseEntity<ApiErrorResponse> handleUnexpectedError(
            Exception exception,
            HttpServletRequest request) {

        log.error(
                "Unexpected error while processing {}",
                request.getRequestURI(),
                exception
        );

        return buildResponse(
                HttpStatus.INTERNAL_SERVER_ERROR,
                "Unexpected server error",
                request.getRequestURI()
        );
    }

    // ============================================================
    // Shared API Error Response
    // ============================================================

    private ResponseEntity<ApiErrorResponse> buildResponse(
            HttpStatus status,
            String message,
            String path) {

        ApiErrorResponse error = new ApiErrorResponse(
                Instant.now(),
                status.value(),
                status.getReasonPhrase(),
                message,
                path
        );

        return ResponseEntity.status(status).body(error);
    }
}
