package de.issa65.order_service.service;

public class AutomationNotFoundException extends RuntimeException {

    public AutomationNotFoundException(String processInstanceKey) {
        super("Process instance not found: " + processInstanceKey);
    }
}