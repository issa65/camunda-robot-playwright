
package de.issa65.order_service.api;

public record OrderResponse(
        String orderNumber,
        String customer,
        Integer quantity,
        String status
) {
}
