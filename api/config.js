export default function handler(request, response) {
  response.setHeader("Cache-Control", "no-store");
  response.status(200).json({ clientId: process.env.GOOGLE_CLIENT_ID || "" });
}