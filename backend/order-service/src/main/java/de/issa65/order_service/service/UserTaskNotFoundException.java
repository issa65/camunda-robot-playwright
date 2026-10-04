package de.issa65.order_service.service;

public class UserTaskNotFoundException extends RuntimeException {

    public UserTaskNotFoundException(String processInstanceKey) {
        super("No active user task found for process instance: "
                + processInstanceKey);
    }
}