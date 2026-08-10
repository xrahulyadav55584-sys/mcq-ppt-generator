import { NextResponse } from "next/server";

export async function POST(req: Request) {
  try {
    const { topic, question } = await req.json();

    const promptText = encodeURIComponent(
      `3D isometric render of ${topic}, ${question}, educational illustration, clean studio lighting, 3D Pixar concept art style`
    );

    const imageUrl = `https://image.pollinations.ai/prompt/${promptText}?width=800&height=800&seed=${Math.floor(
      Math.random() * 1000
    )}&nologo=true`;

    return NextResponse.json({ imageUrl });
  } catch (error) {
    return NextResponse.json({ error: "Failed" }, { status: 500 });
  }
}