const fs = require('fs');
const { Document, Packer, Paragraph, HeadingLevel, TextRun, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle } = require('docx');

const bullet = (text) => new Paragraph({ text, bullet: { level: 0 }, spacing: { after: 120 } });

const sectionTitle = (text) => new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 160 } });
const subTitle = (text) => new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: 220, after: 120 } });

const doc = new Document({
  creator: 'Codex',
  title: 'WhatsApp-Like Chat Application - Project Plan',
  description: 'Node.js project plan for a WhatsApp-style real-time chat platform',
  sections: [
    {
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [
            new TextRun({ text: 'Project Plan Document', bold: true, size: 40 }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 160 },
          children: [new TextRun({ text: 'WhatsApp-Like Chat Application (Node.js)', size: 30 })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [new TextRun({ text: 'Prepared on: February 23, 2026', italics: true })],
        }),

        sectionTitle('1. Project Vision'),
        new Paragraph({ text: 'Build a modern WhatsApp-like chat platform focused on speed, reliability, and polished UX across web and mobile web. The product should support real-time messaging, media sharing, presence indicators, and secure communication with a scalable Node.js backend.', spacing: { after: 140 } }),

        sectionTitle('2. Business Goals'),
        bullet('Launch an MVP in 12 weeks with core 1:1 and group messaging.'),
        bullet('Deliver a premium chat experience comparable to WhatsApp in responsiveness and clarity.'),
        bullet('Support horizontal scaling to at least 100,000 monthly active users in Phase 1.'),
        bullet('Build an architecture that can add voice/video and payments in future phases.'),

        sectionTitle('3. Scope'),
        subTitle('In Scope (MVP)'),
        bullet('User onboarding: phone/email signup, OTP verification, profile setup.'),
        bullet('1:1 chat and group chat with text, emoji, and attachments.'),
        bullet('Message lifecycle: sent, delivered, read receipts, edit, delete.'),
        bullet('Typing indicator and online/offline presence.'),
        bullet('Media upload: images, docs, audio notes, basic compression.'),
        bullet('Search: contacts, chats, and messages.'),
        bullet('Notifications: in-app and push (web push for PWA).'),
        bullet('Admin basics: abuse report handling, user moderation, audit logs.'),

        subTitle('Out of Scope (MVP)'),
        bullet('End-to-end encryption with full cryptographic key rotation (planned Phase 2).'),
        bullet('Video calling and livestream features.'),
        bullet('In-app payments and commerce.'),

        sectionTitle('4. Suggested Tech Stack (Node-Based)'),
        subTitle('Frontend'),
        bullet('Framework: Next.js (React + TypeScript) for SSR and PWA readiness.'),
        bullet('State: Zustand or Redux Toolkit for chat/session state.'),
        bullet('UI: Tailwind CSS + custom design tokens.'),
        bullet('Real-time client transport: Socket.IO client.'),

        subTitle('Backend'),
        bullet('Runtime: Node.js 22+ with TypeScript.'),
        bullet('Framework: NestJS (modular architecture) or Express + clean architecture.'),
        bullet('Real-time: Socket.IO with Redis adapter for multi-instance scale.'),
        bullet('Auth: JWT (access/refresh), OTP service integration.'),
        bullet('Storage: PostgreSQL (core data) + Redis (presence/cache/session).'),
        bullet('Media: S3-compatible object storage + CDN.'),
        bullet('Queue/async: BullMQ with Redis for background jobs.'),

        subTitle('DevOps and Quality'),
        bullet('Containerization: Docker + Docker Compose (dev), Kubernetes optional (prod).'),
        bullet('CI/CD: GitHub Actions for lint, test, build, deploy.'),
        bullet('Observability: OpenTelemetry + Prometheus/Grafana + centralized logs.'),
        bullet('Security: rate limiting, WAF rules, secrets manager, dependency scanning.'),

        sectionTitle('5. System Architecture'),
        new Paragraph({ text: 'Architecture Pattern: Modular monolith in MVP, designed for service extraction later.', spacing: { after: 120 } }),
        bullet('API Gateway Layer: REST + WebSocket handshake/auth.'),
        bullet('Chat Service: manages conversations, messages, receipts, and events.'),
        bullet('Presence Service: online state, last seen, typing state.'),
        bullet('Media Service: upload signing, scanning, storage, CDN URLs.'),
        bullet('Notification Service: push dispatch and retry queues.'),
        bullet('Audit Service: moderation events and compliance logs.'),

        sectionTitle('6. Core Data Model'),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph({ text: 'Entity', bold: true })] }),
              new TableCell({ children: [new Paragraph({ text: 'Purpose', bold: true })] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph('users')] }),
              new TableCell({ children: [new Paragraph('Identity, credentials, profile, privacy settings')] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph('conversations')] }),
              new TableCell({ children: [new Paragraph('1:1/group metadata, ownership, settings')] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph('conversation_members')] }),
              new TableCell({ children: [new Paragraph('Membership, roles, mute/pin/archive state')] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph('messages')] }),
              new TableCell({ children: [new Paragraph('Message content, type, status, edits, deletions')] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph('message_receipts')] }),
              new TableCell({ children: [new Paragraph('Delivered/read timestamps per recipient')] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph('attachments')] }),
              new TableCell({ children: [new Paragraph('File metadata, storage key, scan status')] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph('notifications')] }),
              new TableCell({ children: [new Paragraph('Push queue and delivery outcomes')] }),
            ] }),
          ],
          borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: '999999' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: '999999' },
            left: { style: BorderStyle.SINGLE, size: 1, color: '999999' },
            right: { style: BorderStyle.SINGLE, size: 1, color: '999999' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' },
          },
        }),

        sectionTitle('7. UX and Visual Design Direction'),
        subTitle('Design Principles'),
        bullet('Prioritize message readability first, then feature discoverability.'),
        bullet('Single-thumb mobile ergonomics: actions in reachable zones.'),
        bullet('Fast feedback loops: instant optimistic UI for sending/receiving.'),
        bullet('Consistent visual hierarchy for chat list, conversation pane, and action areas.'),

        subTitle('UI System'),
        bullet('Typography: clean geometric sans-serif with strong weights for names and metadata contrast.'),
        bullet('Color strategy: fresh green accent family with neutral backgrounds and clear unread indicators.'),
        bullet('Spacing: 8-point layout grid; dense but not cramped chat rows.'),
        bullet('Components: reusable chat bubble, message composer, avatar stack, status badges, attachment cards.'),
        bullet('Motion: subtle transitions for new messages, typing indicator pulse, and drawer transitions.'),

        subTitle('Accessibility'),
        bullet('WCAG 2.2 AA contrast targets.'),
        bullet('Keyboard-first support on desktop web.'),
        bullet('Screen-reader labels for message actions and statuses.'),

        sectionTitle('8. Security and Privacy'),
        bullet('Transport security via HTTPS/TLS everywhere.'),
        bullet('Input validation and sanitization for all message/media payloads.'),
        bullet('File scanning pipeline for malware and disallowed file types.'),
        bullet('Role-based admin controls and immutable moderation audit trail.'),
        bullet('Privacy options: last seen visibility, read receipts toggle, block/report controls.'),

        sectionTitle('9. Delivery Roadmap (12 Weeks)'),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph({ text: 'Phase', bold: true })] }),
              new TableCell({ children: [new Paragraph({ text: 'Weeks', bold: true })] }),
              new TableCell({ children: [new Paragraph({ text: 'Key Deliverables', bold: true })] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph('Foundation')] }),
              new TableCell({ children: [new Paragraph('1-2')] }),
              new TableCell({ children: [new Paragraph('Repo setup, CI, auth skeleton, DB schema, UI wireframes')] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph('Core Messaging')] }),
              new TableCell({ children: [new Paragraph('3-5')] }),
              new TableCell({ children: [new Paragraph('Real-time transport, 1:1 chat, receipts, chat list, composer')] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph('Groups and Media')] }),
              new TableCell({ children: [new Paragraph('6-8')] }),
              new TableCell({ children: [new Paragraph('Group chat, attachments, search, notification pipeline')] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph('Hardening')] }),
              new TableCell({ children: [new Paragraph('9-10')] }),
              new TableCell({ children: [new Paragraph('Performance, security, observability, moderation tooling')] }),
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph('UAT and Launch')] }),
              new TableCell({ children: [new Paragraph('11-12')] }),
              new TableCell({ children: [new Paragraph('QA, bug fixes, release prep, production go-live')] }),
            ] }),
          ],
          borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: '999999' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: '999999' },
            left: { style: BorderStyle.SINGLE, size: 1, color: '999999' },
            right: { style: BorderStyle.SINGLE, size: 1, color: '999999' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'cccccc' },
          },
        }),

        sectionTitle('10. Team Composition'),
        bullet('1 Product Manager'),
        bullet('1 UI/UX Designer'),
        bullet('2 Frontend Engineers (React/Next.js)'),
        bullet('2 Backend Engineers (Node.js/NestJS)'),
        bullet('1 QA Engineer (automation + manual)'),
        bullet('1 DevOps Engineer (part-time/shared)'),

        sectionTitle('11. Risks and Mitigation'),
        bullet('High concurrent socket load -> mitigate with Redis adapter, autoscaling, load tests.'),
        bullet('Message delivery edge cases -> idempotent message IDs, robust retry strategy.'),
        bullet('Media abuse/security risks -> strict file policies and scanning workflow.'),
        bullet('Scope creep -> maintain strict MVP feature gate and change control.'),

        sectionTitle('12. Success Metrics'),
        bullet('P95 send-to-deliver latency under 500 ms.'),
        bullet('Crash-free sessions above 99.5%.'),
        bullet('Day-30 retention above 30% for initial cohort.'),
        bullet('Support tickets under defined threshold during first 30 days post-launch.'),

        sectionTitle('13. Deliverables'),
        bullet('Architecture document and API specifications.'),
        bullet('Figma UI kit and finalized high-fidelity screens.'),
        bullet('Production-ready Node.js codebase with CI/CD.'),
        bullet('Deployment runbook and monitoring dashboards.'),
        bullet('MVP launch checklist and post-launch support plan.'),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  const output = 'WhatsApp-Like-Chat-Project-Plan.docx';
  fs.writeFileSync(output, buffer);
  console.log(`Created ${output}`);
});
