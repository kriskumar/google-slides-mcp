# google-slides-mcp
Key Features

Google Slides Integration:

Create new presentations
Add various slide types (title, content, section headers, etc.)
Create tables
Insert images


Data Visualization:

Create various chart types (bar, line, scatter, pie, etc.)
Generate sample data for testing
Customize chart appearance
Automatically add charts to slides

In config.json add mcp server
```
 "google-slides": {
      "command": "python",
      "args": [
        "/path/to/google-slides-server.py"
      ],
      "env": {
        "GOOGLE_TOKEN_PATH": "/path/to/token.json"
      }
    }
```
