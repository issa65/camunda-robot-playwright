
package de.issa65.order_service.service;

import de.issa65.order_service.api.OrderRequest;
import de.issa65.order_service.api.OrderResponse;
import de.issa65.order_service.entity.OrderEntity;
import de.issa65.order_service.repository.OrderRepository;

import io.camunda.client.CamundaClient;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.dao.DataIntegrityViolationException;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.transaction.support.TransactionTemplate;

import java.util.Map;

@Service
public class OrderService {

    private static final Logger log =
            LoggerFactory.getLogger(OrderService.class);

    private final OrderRepository orderRepository;
    private final CamundaClient camundaClient;
    private final TransactionTemplate transactionTemplate;

    public OrderService(
            OrderRepository orderRepository,
            CamundaClient camundaClient,
            TransactionTemplate transactionTemplate) {

        this.orderRepository = orderRepository;
        this.camundaClient = camundaClient;
        this.transactionTemplate = transactionTemplate;
    }

    public OrderResponse createOrder(OrderRequest request) {

        // 1. Doppelte Auftragsnummer prüfen
        if (orderRepository.existsByOrderNumber(
                request.orderNumber())) {
            throw new IllegalArgumentException(
                    "Order already exists: " + request.orderNumber()
            );
        }

        // 2. Auftrag in einer eigenen Transaktion speichern.
        // Nach execute() ist die Transaktion abgeschlossen.
        try {
            transactionTemplate.executeWithoutResult(status -> {
                OrderEntity entity = new OrderEntity(
                        request.orderNumber(),
                        request.customer(),
                        request.quantity()
                );

                orderRepository.saveAndFlush(entity);
            });
        } catch (DataIntegrityViolationException exception) {

            // Auch konkurrierende doppelte POSTs berücksichtigen.
            if (orderRepository.existsByOrderNumber(
                    request.orderNumber())) {
                throw new IllegalArgumentException(
                        "Order already exists: "
                                + request.orderNumber(),
                        exception
                );
            }

            throw exception;
        }

        // 3. Erst nach dem Datenbank-Commit Camunda starten
        try {
            var processInstance = camundaClient
                    .newCreateInstanceCommand()
                    .bpmnProcessId("order-processing")
                    .latestVersion()
                    .variables(Map.of(
                            "orderNumber", request.orderNumber(),
                            "customer", request.customer(),
                            "quantity", request.quantity()
                    ))
                    .send()
                    .join();

            log.info(
                    "Camunda-Prozess {} fuer Auftrag {} gestartet",
                    processInstance.getProcessInstanceKey(),
                    request.orderNumber()
            );

        } catch (Exception exception) {

            log.error(
                    "Camunda-Prozessstart fuer Auftrag {} nicht bestaetigt",
                    request.orderNumber(),
                    exception
            );

            // Auftrag nicht blind löschen:
            // Camunda könnte trotz Timeout gestartet sein.
            throw new IllegalStateException(
                    "Camunda-Prozessstart nicht bestaetigt fuer Auftrag "
                            + request.orderNumber(),
                    exception
            );
        }

        // 4. Antwort auf den erfolgreich angenommenen Auftrag
        return new OrderResponse(
                request.orderNumber(),
                request.customer(),
                request.quantity(),
                "CREATED"
        );
    }

    @Transactional(readOnly = true)
    public OrderResponse getOrder(String orderNumber) {

        return orderRepository.findByOrderNumber(orderNumber)
                .map(this::toResponse)
                .orElse(null);
    }

    @Transactional
    public OrderResponse updateOrderStatus(
            String orderNumber,
            String newStatus) {

        return orderRepository.findByOrderNumber(orderNumber)
                .map(entity -> {
                    entity.updateStatus(newStatus);
                    return toResponse(entity);
                })
                .orElse(null);
    }

    private OrderResponse toResponse(OrderEntity entity) {
        return new OrderResponse(
                entity.getOrderNumber(),
                entity.getCustomer(),
                entity.getQuantity(),
                entity.getStatus()
        );
    }
}
