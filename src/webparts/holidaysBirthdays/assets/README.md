# Default Images

This folder contains default images that are baked into the web part.

## Required Images

- `default-holiday.png` (or `.jpg`) - Default image for generic Holiday events
  - Used as fallback for holidays that don't have specific images
- `default-birthday.png` (or `.jpg`) - Default image for Birthday events
  - Used for all birthday events
- `default-christmas.jpg` - Specific image for Christmas (December 25)
  - Automatically used for events on December 25
- `default-independence.jpg` - Specific image for Independence Day (July 4)
  - Automatically used for events on July 4

## Image Specifications

- **Resolution**: 96x96px to 144x144px (square, 1:1 aspect ratio)
- **File Size**: < 20KB per image (target: 10-15KB)
- **Format**: PNG (with transparency) or JPEG
- **Optimization**: Images should be optimized for web (use tools like TinyPNG, Squoosh, or ImageOptim)

## Usage

These images are automatically used when an event doesn't have a custom `ImageUrl` specified in the SharePoint list.

**Image Selection Logic:**
1. If event has custom `ImageUrl` → use custom image
2. If event is Christmas (December 25) → use `default-christmas.jpg`
3. If event is Independence Day (July 4) → use `default-independence.jpg`
4. If event is Birthday → use `default-birthday.png`
5. If event is other Holiday → use `default-holiday.png`

## Image Files

All images are located in this `assets` folder:

- ✅ `default-christmas.jpg` - Christmas image (December 25)
- ✅ `default-independence.jpg` - Independence Day image (July 4)
- ✅ `default-holiday.jpg` - Generic holiday fallback
- ✅ `default-birthday.jpg` - Birthday image

