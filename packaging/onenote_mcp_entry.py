"""PyInstaller entry script. A plain top-level script (not ``-m``) is the most reliable
freeze entry. It just runs the server's console main; with no args that's the stdio MCP
server, with ``--configure`` it registers the server in Claude Desktop and exits (§8)."""

from onenote_com_mcp.server import main

if __name__ == "__main__":
    main()
