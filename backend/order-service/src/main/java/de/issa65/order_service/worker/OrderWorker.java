
package de.issa65.order_service.worker;

import io.camunda.client.annotation.JobWorker;
import io.camunda.client.api.response.ActivatedJob;
import org.springframework.stereotype.Component;
import de.issa65.order_service.service.OrderService;

import java.util.Map;

@Component
public class OrderWorker {

    private final OrderService orderService;

    public OrderWorker(OrderService orderService) {
        this.orderService = orderService;
    }

    @JobWorker(type = "check-order")
    public Map<String, Object> checkOrder(ActivatedJob job) {

        Map<String, Object> variables = job.getVariablesAsMap();

        String orderNumber =
                String.valueOf(variables.get("orderNumber"));

        Object quantity = variables.get("quantity");

        System.out.println(
                "Prüfe Auftrag: " + orderNumber
                        + ", Menge: " + quantity
        );

        return Map.of(
                "orderChecked", true
        );
    }




    @JobWorker(type = "approve-order")
    public Map<String, Object> approveOrder(ActivatedJob job) {

        Map<String, Object> variables = job.getVariablesAsMap();

        String orderNumber =
                String.valueOf(variables.get("orderNumber"));

        var updatedOrder = orderService.updateOrderStatus(
                orderNumber,
                "COMPLETED"
        );

        if (updatedOrder == null) {
            throw new IllegalStateException(
                    "Auftrag nicht gefunden: " + orderNumber
            );
        }

        System.out.println(
                "Auftrag freigegeben: " + orderNumber
        );

        return Map.of(
                "orderStatus", "COMPLETED"
        );
    }


    @JobWorker(type = "reject-order")
    public Map<String, Object> rejectOrder(ActivatedJob job) {

        Map<String, Object> variables = job.getVariablesAsMap();

        String orderNumber =
                String.valueOf(variables.get("orderNumber"));

        var updatedOrder = orderService.updateOrderStatus(
                orderNumber,
                "REJECTED"
        );

        if (updatedOrder == null) {
            throw new IllegalStateException(
                    "Auftrag nicht gefunden: " + orderNumber
            );
        }

        System.out.println(
                "Auftrag abgelehnt: " + orderNumber
        );

        return Map.of(
                "orderStatus", "REJECTED"
        );
    }

}
