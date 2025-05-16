from PIL import Image, ImageDraw, ImageFont
import os

def create_icon(size, output_path):
    """Create a simple icon with 'DQV' text."""
    # Create a new image with a dark background
    img = Image.new('RGBA', (size, size), (45, 45, 45, 255))
    draw = ImageDraw.Draw(img)
    
    # Try to use a system font, fallback to default
    try:
        font = ImageFont.truetype("arial.ttf", size=int(size/2))
    except:
        font = ImageFont.load_default()
    
    # Draw text
    text = "DQV"
    text_bbox = draw.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]
    
    # Center the text
    x = (size - text_width) / 2
    y = (size - text_height) / 2
    
    # Draw text with a blue color
    draw.text((x, y), text, fill=(0, 120, 212, 255), font=font)
    
    # Save as ICO
    img.save(output_path, format='ICO', sizes=[(size, size)])

def main():
    # Get the AppData path
    app_data = os.path.join(os.getenv('APPDATA'), 'DQV', 'icons')
    os.makedirs(app_data, exist_ok=True)
    
    # Create icons
    create_icon(32, os.path.join(app_data, 'app_icon.ico'))
    create_icon(16, os.path.join(app_data, 'tray_icon.ico'))
    
    print("Icons created successfully!")

if __name__ == "__main__":
    main() 