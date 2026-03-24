import type { Binding } from "../../types";

/**
 * Navigate PowerPoint to the slide containing a binding, then attempt to
 * select the bound shape. Shape selection silently fails on platforms where
 * setSelectedShapes is unavailable — slide navigation always succeeds.
 */
export async function navigateToBinding(binding: Binding): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    const slide = slides.items[binding.slideIndex];
    if (!slide) return;

    slide.setAsActiveSlide();
    await context.sync();

    // Attempt shape selection — not available on all platforms/API versions
    try {
      context.presentation.setSelectedShapes([binding.shapeId]);
      await context.sync();
    } catch {
      // Slide navigation already succeeded; shape selection is best-effort
    }
  });
}
