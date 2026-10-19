# Agent instructions

## Environment variable compatibility

When adding or changing tools, HTTP routes, scripts, authentication, or Graph requests:

- Read the affected package's README configuration sections before implementing the change. For the Microsoft 365 server, use [packages/microsoft365/README.md](packages/microsoft365/README.md#environment-variables) as the configuration contract.
- Apply every relevant documented control to each entry point. Check read-only mode, presets, tool filters, draft requirements, mailbox allowlists, transport restrictions, authentication, API version, extraction limits, and Graph resilience settings where applicable.
- Register tools through `TOOL_METADATA` and `filterTools`. Gate auxiliary HTTP routes with the same allowed-tool set; standalone scripts must enforce applicable restrictions before side effects.
- Reuse the existing authentication configuration, Graph version resolver, and shared resilience layer instead of hardcoding alternatives. Preserve operation safety when deciding which failures can be retried.
- Add regression coverage demonstrating that applicable restrictions block the operation and configured values reach the request. Include alternate entry points affected by the change.
- Update the package README when intentionally changing the configuration contract.
