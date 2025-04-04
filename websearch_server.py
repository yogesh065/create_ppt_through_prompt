# websearch_server.py (MCP Web Search Service)
from mcp.server.fastmcp import FastMCP
from bs4 import BeautifulSoup
import requests
import re
import json



mcp = FastMCP("websearch")

@mcp.tool()
async def web_search(query: str, max_results: int = 5) -> str:
    """Enhanced web search with sanitization and Google parsing"""
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
            "Accept-Language": "en-US,en;q=0.9",
            "Referer": "https://www.google.com/"
        }
        
        # Perform search
        response = requests.get(
            f"https://www.google.com/search?q={query}&num={max_results}",
            headers=headers,
            timeout=15
        )
        response.raise_for_status()
        
        # Parse results
        soup = BeautifulSoup(response.text, 'html.parser')
        results = []
        
        # Find all search result blocks
        for result in soup.find_all('div', class_='tF2Cxc'):
            link = result.find('a')['href']
            title = result.find('h3', class_='LC20lb').text
            snippet = result.find('div', class_='VwiC3b').text if result.find('div', class_='VwiC3b') else ''
            
            # Validate and sanitize URL
            if re.match(r'^https?://', link):
                results.append({
                    "title": title.strip(),
                    "url": link.split('&')[0],  # Remove tracking parameters
                    "snippet": snippet.strip()
                })
                
            if len(results) >= max_results:
                break
                
        return json.dumps({
            "query": query,
            "results": results
        }, ensure_ascii=False)
        
    except Exception as e:
        return json.dumps({
            "error": str(e),
            "query": query
        }, ensure_ascii=False)

if __name__ == "__main__":
    mcp.run()
