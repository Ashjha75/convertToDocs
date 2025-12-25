# ==============================================================================
# 📘 PPT GENERATOR - JSON STRUCTURE GUIDE
# ==============================================================================
# This file documents the exact JSON format expected by the PPT Generator.
# Use this as a reference when building your input data.
#
# ⚠️ IMPORTANT RULES:
# 1. The JSON must have 3 main sections: "meta", "theme", and "slides".
# 2. Colors are always [R, G, B] lists (0-255).
# 3. "type" fields are case-sensitive (e.g., use "cover", not "Cover").
# ==============================================================================

reference_json = {
  
  # 1. META INFORMATION (Required)
  # ------------------------------
  "meta": {
    "title": "Project Name",      # String: Appears in footer
    "author": "Author Name",      # String: Internal metadata
    "date": "Q4 2025"             # String: Appears in footer
  },

  # 2. THEME CONFIGURATION (Optional - Defaults will be used if omitted)
  # ------------------------------------------------------------------
  "theme": {
    "colors": {
      # [Red, Green, Blue] - Values must be between 0 and 255
      "background": [255, 255, 255],  # Slide background color
      "title": [30, 60, 114],         # Main headings color
      "text": [42, 42, 42],           # Body text color
      "accent": [42, 82, 152],        # Decorative lines, banners, bars
      "table_header": [30, 60, 114],  # Background of table headers
      "border": [200, 200, 200]       # Card borders and grid lines
    },
    "fonts": {
      "family": "Arial",              # Font Name (Must be installed on system)
      "size_title": 32,               # Integer: Title font size (Rec: 28-44)
      "size_body": 16                 # Integer: Body text size (Rec: 12-20)
    }
  },

  # 3. SLIDES (List of Slide Objects)
  # ---------------------------------
  "slides": [
    
    # TYPE: COVER
    # Used for: The very first slide of the presentation.
    {
      "type": "cover",
      "title": "Main Title Here",     # Max ~50 chars
      "subtitle": "Subtitle Here"     # Optional. Max ~100 chars
    },

    # TYPE: CONTENT
    # Used for: Standard text slides, bullet points, summaries.
    {
      "type": "content",
      "title": "Slide Title",
      "banner": "Key Takeaway",       # Optional: Highlighted box at top. Max ~150 chars.
      "text": "Line 1\nLine 2"         # Use '\n' for new lines.
    },

    # TYPE: CARDS
    # Used for: Lists of grouped items. Automatically arranges in a grid.
    # LIMITS: Best with 2, 3, or 4 cards. More than 4 may look crowded.
    {
      "type": "cards",
      "title": "Key Topics",
      "cards": [
        {
          "title": "Card Title",
          "items": [                  # List of strings (Bullet points)
            "Point 1",
            "Point 2"
          ]
        }
      ]
    },

    # TYPE: TABLE
    # Used for: Data, budgets, status lists.
    # LIMITS: Max ~10 columns, Max ~12 rows (otherwise it overflows).
    {
      "type": "table",
      "title": "Budget",
      "columns": [
        { "key": "col1", "label": "Item Name" },  # 'key' must match row keys
        { "key": "col2", "label": "Cost" }
      ],
      "rows": [
        { 
          "col1": "Server", 
          "col2": "$500",
          "status": "Approved"        # SPECIAL KEY: 'status'
                                      # Values: "Approved", "Completed", "On Track" -> GREEN
                                      # Values: "Pending", "Delayed", "At Risk" -> RED
        }
      ]
    },

    # TYPE: CHART
    # Used for: Bar, Column, or Line charts.
    {
      "type": "chart",
      "title": "Growth",
      "chart_type": "COLUMN",         # Options: "COLUMN", "BAR", "LINE"
      "chart": {
        "categories": ["Q1", "Q2"],   # X-Axis Labels
        "series_name": "Revenue",     # Legend Label
        "values": [100.5, 200.0]      # Y-Axis Values (Numbers only)
      }
    },

    # TYPE: GANTT
    # Used for: Project timelines.
    # LIMITS: Max ~10 tasks.
    {
      "type": "gantt",
      "title": "Roadmap",
      "tasks": [
        {
          "name": "Phase 1",          # Task Name
          "start_week": 1,            # Integer: Start week number
          "duration": 4,              # Integer: Duration in weeks
          "progress": 50              # Integer: 0-100 (Percentage)
        }
      ]
    },

    # TYPE: IMAGE
    # Used for: Full slide images/diagrams.
    {
      "type": "image",
      "title": "Architecture",
      "image_path": "diagram.png",    # Path to file. Must exist in folder.
      "caption": "Figure 1"           # Optional caption below image.
    }
  ]
}
