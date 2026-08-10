import { NextResponse } from "next/server";

export async function POST(req: Request) {
  try {
    const { topic, question } = await req.json();

    const prompt = `3D isometric render of ${topic}, related to ${question}, educational illustration, clean studio lighting, 3D Pixar concept art style, high detail, isolated background`;

    const encodedPrompt = encodeURIComponent(prompt);
    const imageUrl = `https://image.pollinations.ai/prompt/${encodedPrompt}?width=800&height=800&seed=42&nologo=true`;

    return NextResponse.json({ imageUrl });
  } catch (error) {
    return NextResponse.json({ error: "Failed to generate image" }, { status: 500 });
  }
}