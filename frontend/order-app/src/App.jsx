
import { useState } from "react";

function App() {
  const [orderNumber, setOrderNumber] = useState("");
  const [customer, setCustomer] = useState("");
  const [quantity, setQuantity] = useState(1);
  const [createdOrder, setCreatedOrder] = useState(null);


async function handleSubmit(event) {
  event.preventDefault();

  setCreatedOrder(null);

  try {
    const response = await fetch("/api/orders", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        orderNumber,
        customer,
        quantity: Number(quantity),
      }),
    });

    if (!response.ok) {
      throw new Error(
        `Auftrag konnte nicht erstellt werden (HTTP ${response.status})`
      );
    }

    const savedOrder = await response.json();

    setCreatedOrder(savedOrder);

  } catch (error) {
    console.error("Fehler beim Anlegen:", error);
    alert(error.message);
  }
}


  return (
    <main style={{
      maxWidth: 600,
      margin: "60px auto",
      fontFamily: "Arial, sans-serif",
      padding: 24,
    }}>
      <h1>Order Management – E2E Lab</h1>
      <p>Neuen Auftrag anlegen</p>

      <form onSubmit={handleSubmit}>
        <div style={{ marginBottom: 16 }}>
          <label htmlFor="orderNumber">Auftragsnummer</label>
          <input
            id="orderNumber"
            data-testid="order-number"
            value={orderNumber}
            onChange={(e) => setOrderNumber(e.target.value)}
            required
            style={{ display: "block", width: "100%", padding: 10 }}
          />
        </div>

        <div style={{ marginBottom: 16 }}>
          <label htmlFor="customer">Kunde</label>
          <input
            id="customer"
            data-testid="customer"
            value={customer}
            onChange={(e) => setCustomer(e.target.value)}
            required
            style={{ display: "block", width: "100%", padding: 10 }}
          />
        </div>

        <div style={{ marginBottom: 16 }}>
          <label htmlFor="quantity">Menge</label>
          <input
            id="quantity"
            data-testid="quantity"
            type="number"
            min="1"
            value={quantity}
            onChange={(e) => setQuantity(Number(e.target.value))}
            required
            style={{ display: "block", width: "100%", padding: 10 }}
          />
        </div>

        <button
          type="submit"
          data-testid="submit-order"
          style={{ padding: "12px 20px", cursor: "pointer" }}
        >
          Auftrag anlegen
        </button>
      </form>

      {createdOrder && (
        <section
          data-testid="order-success"
          style={{
            marginTop: 24,
            padding: 16,
            background: "#e5f7e9",
            color: "#14532d",
            borderRadius: 8,
          }}
        >
          <h2>Auftrag erfolgreich angelegt!</h2>
          <p>Auftragsnummer: {createdOrder.orderNumber}</p>
          <p>Kunde: {createdOrder.customer}</p>
          <p>Menge: {createdOrder.quantity}</p>
        </section>
      )}
    </main>
  );
}

export default App;