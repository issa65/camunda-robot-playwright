
package de.issa65.order_service.api;

import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.NotNull;

public record OrderRequest(

        @NotBlank
        String orderNumber,

        @NotBlank
        String customer,

        @NotNull
        @Min(1)
        Integer quantity

) {
}
