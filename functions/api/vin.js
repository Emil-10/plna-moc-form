const DK_API_URL = "https://api.dataovozidlech.cz/api/vehicletechnicaldata/v2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type"
};

export async function onRequestOptions() {
  return new Response(null, { status: 204, headers: CORS_HEADERS });
}

export async function onRequestGet(context) {
  const url = new URL(context.request.url);
  const vin = (url.searchParams.get("vin") || "").trim().toUpperCase();

  if (!vin || vin.length !== 17) {
    return jsonResponse({ error: "Parametr vin musí mít 17 znaků." }, 400);
  }

  const apiKey = context.env.DK_API_KEY;
  if (!apiKey) {
    return jsonResponse({ error: "API klíč není nastaven (DK_API_KEY)." }, 500);
  }

  try {
    const response = await fetch(`${DK_API_URL}?vin=${encodeURIComponent(vin)}`, {
      headers: { "API_KEY": apiKey }
    });

    const data = await response.json();
    return jsonResponse(data, response.ok ? 200 : response.status);
  } catch (err) {
    return jsonResponse({ error: "Nepodařilo se spojit s registrem vozidel." }, 502);
  }
}

function jsonResponse(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...CORS_HEADERS }
  });
}
