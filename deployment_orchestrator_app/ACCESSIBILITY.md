# Accessibility readiness

Windows Audio and Display Baseline Enforcer Orchestrator targets WCAG 2.2 Level AA and the applicable software criteria in Revised Section 508. This document records engineering support and remaining validation work; it is not an Accessibility Conformance Report (ACR) and does not by itself establish conformance.

## Implemented support

- Standard Flutter Material controls provide keyboard activation, focus traversal, control names, roles, and states.
- A reading-order focus traversal group covers the unified workspace.
- Light, dark, and operating-system high-contrast themes are provided.
- Text fields use a two-pixel high-contrast focused border; buttons and custom result tiles retain visible focus treatment.
- Standard actions are at least 44–48 logical pixels tall. Compact per-PC icon actions are 40 by 40 logical pixels and include descriptive tooltips.
- Dynamic deployment and monitoring statuses are exposed as live semantic regions.
- Progress indicators expose a name and human-readable value to assistive technology.
- Truncated computer names retain their complete name in tooltips and semantic labels.
- Color is never the only status indicator: status text and distinct icons accompany every status color.
- Status colors use separate light- and dark-theme values selected for readable foreground contrast.
- The primary workspace supports scrolling and a persistent desktop scrollbar.
- The app is exercised at 200 percent text scaling by an automated widget test.

## Required validation before publishing an ACR/VPAT

An ACR should be based on test evidence for the exact packaged Windows release. At minimum:

1. Complete every workflow using only the keyboard, including dialogs, dropdowns, autocomplete results, PC detail tiles, retries, and stopping a process.
2. Test with Windows Narrator and a current NVDA release. Confirm names, roles, states, reading order, live status announcements, and dialog focus behavior.
3. Test Windows Contrast Themes and both app color modes, including disabled controls and all deployment/monitoring statuses.
4. Test Windows text scaling and display scaling through 200 percent at the minimum supported window size. Confirm no information or operation is lost.
5. Measure text, icon, control-boundary, and focus-indicator contrast with an accessibility contrast tool.
6. Verify all pointer targets, error messages, time-dependent behavior, and process cancellation behavior.
7. Document defects and remediation, then complete the applicable WCAG Level A/AA and Revised Section 508 software tables using only the standard conformance terms.

Automated Flutter tests are useful regression checks, but they do not replace assistive-technology and human usability testing.
