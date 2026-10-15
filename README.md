# I01 Outlook probe pane (static, feasibility testing only)

Static task-pane files for the read-only Outlook feasibility probe of the private project
`yinquan251/quanyin-workbuddy-mail-expert` (planned issue I01). Published only because Outlook
add-in panes must be served from an HTTPS origin.

- Contains no secrets, no mail, no server code. The pane reports metadata only and never sends,
  saves, or changes mail.
- Its only network calls go to `127.0.0.1` / `localhost` probe servers on the tester's own machine.
- This repository will be deleted after the I01 feasibility tests.
