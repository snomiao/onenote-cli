# Microsoft Graph OneNote API Endpoints

Reference for the OneNote REST API endpoints used by `onenote-cli`.

## Base URL

```
https://graph.microsoft.com/v1.0/me/onenote/
```

## Notebooks

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/me/onenote/notebooks` | List all notebooks |
| GET | `/me/onenote/notebooks/{id}` | Get a notebook by ID |
| POST | `/me/onenote/notebooks` | Create a notebook (`{ "displayName": "..." }`) |

## Sections

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/me/onenote/sections` | List all sections |
| GET | `/me/onenote/notebooks/{id}/sections` | List sections in a notebook |
| GET | `/me/onenote/sections/{id}` | Get a section by ID |
| POST | `/me/onenote/notebooks/{id}/sections` | Create a section (`{ "displayName": "..." }`) |

## Section Groups

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/me/onenote/sectionGroups` | List all section groups |
| GET | `/me/onenote/notebooks/{id}/sectionGroups` | List section groups in a notebook |

## Pages

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/me/onenote/pages` | List all pages |
| GET | `/me/onenote/sections/{id}/pages` | List pages in a section |
| GET | `/me/onenote/pages/{id}` | Get page metadata |
| GET | `/me/onenote/pages/{id}/content` | Get page HTML content |
| POST | `/me/onenote/sections/{id}/pages` | Create a page (Content-Type: text/html) |
| DELETE | `/me/onenote/pages/{id}` | Delete a page |

## Search

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/me/onenote/pages?$search={query}` | Search pages by content |

## Creating Pages

Pages are created by POSTing HTML to a section:

```http
POST /me/onenote/sections/{section-id}/pages
Content-Type: text/html

<!DOCTYPE html>
<html>
  <head>
    <title>Page Title</title>
  </head>
  <body>
    <p>Page content here</p>
  </body>
</html>
```

## Required Permissions (Delegated)

| Permission | Description |
|------------|-------------|
| `Notes.Read` | Read notebooks |
| `Notes.ReadWrite` | Read and write notebooks |
| `Notes.Read.All` | Read all accessible notebooks |
| `Notes.ReadWrite.All` | Read and write all accessible notebooks |

## Additional Scopes

Notebooks can also be accessed via group or SharePoint site contexts:

```
/groups/{id}/onenote/{notebooks|sections|pages}
/sites/{id}/onenote/{notebooks|sections|pages}
```

These require corresponding group/site read permissions.

## References

- [OneNote API overview](https://learn.microsoft.com/en-us/graph/integrate-with-onenote)
- [OneNote REST API reference](https://learn.microsoft.com/en-us/graph/api/resources/onenote-api-overview)
- [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
