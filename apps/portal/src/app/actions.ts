"use server";

import { createLicenseKey } from "../lib/crypto";

// A reason the buyer can act on; anything else is logged and shown as a generic error
class Refusal extends Error {}

const ORDER_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export async function validateOrderAndGenerateLicense(orderId: string) {
  try {
    // Polar order IDs are UUIDs; checking the shape keeps the URL below to exactly one order
    if (typeof orderId !== "string" || !ORDER_ID.test(orderId.trim())) {
      throw new Refusal("That doesn't look like a Polar order ID. Copy it from your receipt.");
    }
    orderId = orderId.trim().toLowerCase();

    const token = process.env.POLAR_ACCESS_TOKEN;
    const expectedProductId = process.env.POLAR_PRODUCT_ID;
    if (!token || !expectedProductId) {
      throw new Error("POLAR_ACCESS_TOKEN or POLAR_PRODUCT_ID is missing on the server.");
    }
    // Use the custom POLAR_API_URL if provided, else default to production.
    const apiUrl = process.env.POLAR_API_URL || "https://api.polar.sh/v1/orders";
    // Sandbox orders are free, so a live site must never accept them
    if (process.env.NODE_ENV === "production" && apiUrl.includes("sandbox") && process.env.ALLOW_SANDBOX !== "1") {
      throw new Error("POLAR_API_URL points at the Polar sandbox in production. Set ALLOW_SANDBOX=1 only for a test deployment.");
    }

    const res = await fetch(`${apiUrl}/${orderId}`, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    if (!res.ok) {
      if (res.status === 404 || res.status === 422) {
        throw new Refusal("Order not found or invalid.");
      }
      throw new Error(`Polar API error: ${res.status} ${res.statusText}`);
    }

    const orderData = await res.json();

    // The product decides the tier, not the amount: PPP discounts make the price vary by country
    if (orderData.product_id !== expectedProductId) {
      throw new Refusal("This order is not for XlsxFlow Pro.");
    }
    if (orderData.status !== "paid") {
      throw new Refusal(`This order is ${typeof orderData.status === "string" ? orderData.status.replace(/_/g, " ") : "not paid"}.`);
    }

    return { success: true, license: createLicenseKey(orderId) };
  } catch (error) {
    if (error instanceof Refusal) return { success: false, error: error.message };
    console.error("License request failed:", error);
    return { success: false, error: "Something went wrong on our side. Please try again later." };
  }
}
