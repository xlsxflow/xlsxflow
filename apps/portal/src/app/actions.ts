"use server";

import { createLicenseKey } from "../lib/crypto";

export async function validateOrderAndGenerateLicense(orderId: string) {
  try {
    const token = process.env.POLAR_ACCESS_TOKEN;
    if (!token) {
      throw new Error("Polar Access Token is missing on the server.");
    }

    // Call the Polar.sh API to verify the order.
    // Use the custom POLAR_API_URL if provided, else default to production.
    const apiUrl = process.env.POLAR_API_URL || "https://api.polar.sh/api/v1/orders";
    const res = await fetch(`${apiUrl}/${encodeURIComponent(orderId)}`, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    if (!res.ok) {
      if (res.status === 404) {
        throw new Error("Order not found or invalid.");
      }
      throw new Error(`Polar API error: ${res.statusText}`);
    }

    const orderData = await res.json();

    // The product decides the tier, not the amount: PPP discounts make the price vary by country
    const expectedProductId = process.env.POLAR_PRODUCT_ID;
    if (!expectedProductId) {
      throw new Error("POLAR_PRODUCT_ID is missing on the server.");
    }
    if (orderData.product_id !== expectedProductId) {
      throw new Error("Invalid Product: This order ID is not for the SheetForge Pro license.");
    }
    if (orderData.status !== "paid") {
      throw new Error(`Order is ${orderData.status ?? "not paid"}.`);
    }

    return { success: true, license: createLicenseKey(orderId) };
  } catch (error: any) {
    return { success: false, error: error.message };
  }
}
