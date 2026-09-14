"use server";

import { createLicenseJson } from "../lib/crypto";

export async function validateOrderAndGenerateLicense(orderId: string, hardwareHash: string) {
  try {
    const token = process.env.POLAR_ACCESS_TOKEN;
    if (!token) {
      throw new Error("Polar Access Token is missing on the server.");
    }

    // Call the Polar.sh API to verify the order.
    // Use the custom POLAR_API_URL if provided, else default to production.
    const apiUrl = process.env.POLAR_API_URL || "https://api.polar.sh/api/v1/orders";
    const res = await fetch(`${apiUrl}/${orderId}`, {
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
    
    // STRICT SECURITY VALIDATION
    // 1. Validate Product ID to prevent users from buying a cheap $1 product and using its Order ID for the $5 Pro Tier
    const expectedProductId = process.env.POLAR_PRODUCT_ID;
    if (expectedProductId && orderData.product_id !== expectedProductId) {
      throw new Error("Invalid Product: This order ID is not for the SheetForge Pro license.");
    }

    // 2. Validate Order Amount as a fallback if product ID is missing in dev
    const expectedAmount = 500; // $5.00 in cents
    if (orderData.amount < expectedAmount) {
      throw new Error(`Order amount is too low. Expected at least $5.00, got $${(orderData.amount / 100).toFixed(2)}`);
    }

    // Generate the license using the secure crypto logic
    const license = await createLicenseJson(orderId, hardwareHash);
    
    return { success: true, license: JSON.stringify(license, null, 2), data: orderData };
  } catch (error: any) {
    return { success: false, error: error.message };
  }
}
