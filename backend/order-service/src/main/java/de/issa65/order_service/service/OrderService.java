
package de.issa65.order_service.service;

import de.issa65.order_service.api.OrderRequest;
import de.issa65.order_service.api.OrderResponse;
import org.springframework.stereotype.Service;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import io.camunda.client.CamundaClient;


@Service
public class OrderService {

    private final Map<String, OrderResponse> orders =
            new ConcurrentHashMap<>();

    private final CamundaClient camundaClient;

    public OrderService(CamundaClient camundaClient) {
        this.camundaClient = camundaClient;
    }
    
    public OrderResponse createOrder(OrderRequest request) {

        OrderResponse order = new OrderResponse(
                request.orderNumber(),
                request.customer(),
                request.quantity(),
                "CREATED"
        );

        OrderResponse existing = orders.putIfAbsent(
                request.orderNumber(),
                order
        );

        if (existing != null) {
            throw new IllegalArgumentException(
                    "Order already exists: " + request.orderNumber()
            );
        }

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

            System.out.println(
                    "Camunda-Prozess gestartet: "
                            + processInstance.getProcessInstanceKey()
                            + " für Auftrag " + request.orderNumber()
            );

        } catch (Exception exception) {

            // Auftrag zurücknehmen, falls Camunda nicht startet.
            orders.remove(request.orderNumber(), order);

            throw new IllegalStateException(
                    "Camunda-Prozess konnte nicht gestartet werden "
                            + "für Auftrag " + request.orderNumber(),
                    exception
            );
        }

        return order;
    }

    public OrderResponse getOrder(String orderNumber) {
        return orders.get(orderNumber);
    }

    public OrderResponse updateOrderStatus(
            String orderNumber,
            String newStatus) {

        return orders.computeIfPresent(
                orderNumber,
                (key, existing) -> new OrderResponse(
                        existing.orderNumber(),
                        existing.customer(),
                        existing.quantity(),
                        newStatus
                )
        );
    }
}
