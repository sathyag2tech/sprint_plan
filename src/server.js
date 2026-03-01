const path = require("path");
const express = require("express");
const { Document, Packer, Paragraph, HeadingLevel, TextRun, AlignmentType } = require("docx");

const PORT = process.env.PORT || 3000;
const app = express();

app.use(express.json({ limit: "1mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "..", "public")));

function clean(value, max = 1000) {
  return String(value || "").trim().slice(0, max);
}

function splitList(value) {
  return clean(value, 3000)
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function line(text) {
  return new Paragraph({ text, spacing: { after: 120 } });
}

function bullets(items) {
  if (!items.length) {
    return [line("- Not specified")];
  }

  return items.map((item) =>
    new Paragraph({
      text: item,
      bullet: { level: 0 },
      spacing: { after: 80 },
    })
  );
}

function section(title) {
  return new Paragraph({
    text: title,
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 260, after: 120 },
  });
}

function sub(title) {
  return new Paragraph({
    text: title,
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 180, after: 100 },
  });
}

function parseTimelineWeeks(timelineText) {
  const text = clean(timelineText, 120).toLowerCase();
  if (!text) return 12;

  const weekMatch = text.match(/(\d+)\s*week/);
  if (weekMatch) return Math.min(Math.max(Number(weekMatch[1]), 4), 24);

  const monthMatch = text.match(/(\d+)\s*month/);
  if (monthMatch) return Math.min(Math.max(Number(monthMatch[1]) * 4, 4), 24);

  const numericMatch = text.match(/\d+/);
  if (numericMatch) return Math.min(Math.max(Number(numericMatch[0]), 4), 24);

  return 12;
}

function toLowerList(items) {
  return items.map((item) => item.toLowerCase());
}

function buildArchitectureRecommendations(coreFeatures, integrations, securityNeeds, techStack) {
  const featureSignals = toLowerList(coreFeatures);
  const integrationSignals = toLowerList(integrations);
  const securityText = securityNeeds.toLowerCase();
  const stackText = techStack.toLowerCase();

  const architecture = [
    "Client layer: responsive web app with component-driven UI and centralized state management.",
    "API layer: Node.js backend with REST APIs for onboarding, profile, contacts, and admin operations.",
    "Realtime layer: Socket.IO gateway for message delivery, typing events, presence, and read receipts.",
    "Data layer: PostgreSQL for durable relational data and Redis for session, cache, and pub/sub events.",
    "Async processing: job queue for notifications, media processing, retries, and background automation.",
  ];

  if (featureSignals.some((item) => item.includes("group"))) {
    architecture.push("Conversation domain split into direct chat and group chat policies with role controls.");
  }

  if (featureSignals.some((item) => item.includes("media") || item.includes("file") || item.includes("image"))) {
    architecture.push("Media pipeline with object storage, signed URLs, metadata store, and virus scanning hook.");
  }

  if (integrationSignals.some((item) => item.includes("firebase") || item.includes("fcm"))) {
    architecture.push("Push notification adapter for Firebase Cloud Messaging with retry and dead-letter strategy.");
  }

  if (integrationSignals.some((item) => item.includes("twilio"))) {
    architecture.push("Communication integration module for OTP/SMS verification and delivery status tracking.");
  }

  if (securityText.includes("e2e") || securityText.includes("encryption")) {
    architecture.push("Security extension: key lifecycle service for end-to-end encryption rollout in controlled phases.");
  }

  if (stackText.includes("nest")) {
    architecture.push("Recommended implementation style: modular monolith with NestJS feature modules.");
  } else {
    architecture.push("Recommended implementation style: Express modular architecture with strict domain boundaries.");
  }

  return architecture;
}

function buildModulePlan(coreFeatures) {
  const featureSignals = toLowerList(coreFeatures);
  const modules = [
    "Auth and Identity: signup/login, token lifecycle, session management, and profile setup.",
    "User Graph: contacts, blocked users, privacy settings, and discoverability options.",
    "Conversation Engine: create/list conversations, participant management, and conversation metadata.",
    "Message Engine: create/edit/delete messages, delivery states, read receipts, and idempotency controls.",
    "Presence and Typing: online status, last-seen, typing signals, and socket state synchronization.",
    "Notification Service: in-app + push triggers with user preference filters and retry logic.",
    "Observability: structured logs, trace IDs, metrics, health probes, and alert thresholds.",
  ];

  if (featureSignals.some((item) => item.includes("admin") || item.includes("moderation"))) {
    modules.push("Admin and Moderation: report workflow, user suspension, and audit logging.");
  }

  if (featureSignals.some((item) => item.includes("search"))) {
    modules.push("Search Module: indexed retrieval for users, conversations, and messages.");
  }

  if (featureSignals.some((item) => item.includes("media") || item.includes("attachment") || item.includes("file"))) {
    modules.push("Media Module: upload authorization, attachment metadata, file lifecycle, and retention policies.");
  }

  return modules;
}

function buildDataModelPlan(coreFeatures) {
  const featureSignals = toLowerList(coreFeatures);
  const entities = [
    "users: account identity, profile metadata, privacy controls, and authentication references.",
    "devices: user device tokens, platform info, and session state for push delivery routing.",
    "conversations: direct/group conversation metadata, configuration, and archival status.",
    "conversation_members: role, mute/pin status, join/leave timestamps, and permission flags.",
    "messages: message body, type, author, client message ID, and lifecycle timestamps.",
    "message_receipts: per-recipient delivered/read timestamps and read-state transitions.",
    "presence_events: online/offline transitions and activity snapshots.",
  ];

  if (featureSignals.some((item) => item.includes("media") || item.includes("attachment") || item.includes("file"))) {
    entities.push("attachments: storage path, MIME type, file metadata, scan status, and ownership mapping.");
  }

  if (featureSignals.some((item) => item.includes("notification") || item.includes("push"))) {
    entities.push("notification_jobs: queued notification payloads, provider responses, and retry attempts.");
  }

  return entities;
}

function buildApiPlan() {
  return [
    "Auth APIs: /auth/register, /auth/login, /auth/refresh, /auth/logout.",
    "User APIs: /users/me, /users/:id, /users/:id/privacy, /contacts.",
    "Conversation APIs: /conversations, /conversations/:id, /conversations/:id/members.",
    "Message APIs: /conversations/:id/messages, /messages/:id, /messages/:id/read.",
    "Socket events: message.send, message.received, message.read, typing.start, typing.stop, presence.update.",
    "Admin APIs: /reports, /moderation/actions, /audit/logs.",
  ];
}

function buildExecutiveSummary(summary, targetUsers, coreFeatures, platforms) {
  const featureSignals = toLowerList(coreFeatures);
  const platformSignals = toLowerList(platforms);
  const userText = targetUsers.toLowerCase();
  const summaryText = summary.toLowerCase();

  const lines = [
    "The product is positioned as a realtime communication platform with a Node.js-first architecture and delivery focus on reliability, scalability, and user retention.",
  ];

  if (featureSignals.some((x) => x.includes("group"))) {
    lines.push("The roadmap should prioritize both direct and group conversation experiences to support collaborative usage patterns.");
  }
  if (featureSignals.some((x) => x.includes("media") || x.includes("attachment") || x.includes("file"))) {
    lines.push("Media and attachment handling should be treated as a first-class capability with storage, security, and processing pipelines.");
  }
  if (userText.includes("business") || userText.includes("enterprise")) {
    lines.push("The target segment indicates stronger requirements for governance, auditability, and role-based controls.");
  } else if (userText.includes("student") || userText.includes("community")) {
    lines.push("The target segment suggests emphasizing onboarding simplicity, mobile usability, and engagement-oriented UX.");
  }
  if (platformSignals.some((x) => x.includes("ios") || x.includes("android") || x.includes("mobile"))) {
    lines.push("Execution should remain mobile-first in interaction design while preserving desktop operational workflows.");
  }
  if (summaryText.includes("whatsapp")) {
    lines.push("Benchmark expectations align with WhatsApp-like speed, message clarity, and low-friction interactions.");
  }

  return lines;
}

function buildAudienceAndPlatformPlan(targetUsers, platforms) {
  const userText = targetUsers.toLowerCase();
  const platformSignals = toLowerList(platforms);

  const out = [
    "Primary audience profile should be defined through onboarding assumptions, expected chat frequency, and privacy expectations.",
    "Platform rollout should begin with the highest traffic channel and maintain shared feature parity for core messaging behavior.",
  ];

  if (platformSignals.some((x) => x.includes("web"))) {
    out.push("Web delivery should include responsive breakpoints, keyboard-first navigation, and session persistence controls.");
  }
  if (platformSignals.some((x) => x.includes("android") || x.includes("ios") || x.includes("mobile"))) {
    out.push("Mobile experience should optimize thumb-zone actions, realtime feedback states, and low-bandwidth behavior.");
  }
  if (userText.includes("admin") || userText.includes("support") || userText.includes("team")) {
    out.push("Operational personas require moderation and support tooling from the MVP phase.");
  }

  return out;
}

function buildDeliveryConstraintPlan(timelineWeeks, budget, teamSize) {
  const budgetText = budget.toLowerCase();
  const teamMatch = teamSize.match(/\d+/);
  const teamNumber = teamMatch ? Number(teamMatch[0]) : null;

  const out = [
    `Execution baseline is structured for a ${timelineWeeks}-week delivery window with phased quality gates.`,
    "Scope governance should separate MVP-critical capabilities from deferred enhancements starting in week 2.",
  ];

  if (timelineWeeks <= 8) {
    out.push("Compressed schedule requires strict feature gating and earlier QA automation to avoid release instability.");
  } else if (timelineWeeks >= 14) {
    out.push("Extended timeline allows parallel hardening tracks for observability, security, and release automation.");
  }

  if (teamNumber && teamNumber <= 4) {
    out.push("Team capacity implies serialized feature sequencing; parallel streams should be limited to reduce coordination overhead.");
  } else if (teamNumber && teamNumber >= 8) {
    out.push("Larger team capacity supports parallel domain squads with explicit API contracts and integration checkpoints.");
  }

  if (budgetText.includes("low") || budgetText.includes("tight")) {
    out.push("Cost constraints favor managed infrastructure components and minimal custom platform engineering in MVP.");
  } else {
    out.push("Budget planning should reserve contingency for load testing, security validation, and post-launch hypercare.");
  }

  return out;
}

function buildFeatureScopePlan(coreFeatures, advancedFeatures) {
  const featureSignals = toLowerList(coreFeatures);
  const outCore = [
    "MVP scope should lock around authenticated messaging reliability, conversation management, and delivery-state consistency.",
  ];
  const outAdvanced = [
    "Phase 2 should prioritize differentiators only after stability, observability, and support workflows are production-ready.",
  ];

  if (featureSignals.some((x) => x.includes("group"))) {
    outCore.push("Group conversation controls should include membership roles, moderation constraints, and permission-aware actions.");
  }
  if (featureSignals.some((x) => x.includes("search"))) {
    outCore.push("Search capability should cover users, conversations, and message history with pagination and indexing strategy.");
  }
  if (featureSignals.some((x) => x.includes("media") || x.includes("attachment") || x.includes("file"))) {
    outCore.push("Attachment support should enforce file policy, upload lifecycle, and secure retrieval controls.");
  }
  if (featureSignals.some((x) => x.includes("read") || x.includes("receipt"))) {
    outCore.push("Delivery and read receipts should be modeled with per-recipient state transitions and idempotent updates.");
  }

  const advancedSignals = toLowerList(advancedFeatures);
  if (advancedSignals.some((x) => x.includes("video") || x.includes("call"))) {
    outAdvanced.push("Calling features should be delivered in a dedicated phase due to signaling, QoS, and infrastructure complexity.");
  }
  if (advancedSignals.some((x) => x.includes("ai") || x.includes("assistant"))) {
    outAdvanced.push("AI enhancements should be introduced after baseline telemetry proves stable message and retention patterns.");
  }

  return { outCore, outAdvanced };
}

function buildIntegrationStrategy(integrations) {
  const signals = toLowerList(integrations);
  const out = [
    "Integrations should be isolated behind provider abstraction interfaces to minimize vendor lock-in and simplify failover handling.",
    "Each provider dependency should include retry policy, observability hooks, and quota/health monitoring.",
  ];

  if (signals.some((x) => x.includes("twilio"))) {
    out.push("OTP/SMS delivery should use Twilio-backed verification flow with rate controls and fallback messaging paths.");
  }
  if (signals.some((x) => x.includes("firebase") || x.includes("fcm"))) {
    out.push("Push notification orchestration should use FCM topic/token strategy with preference-aware dispatch.");
  }
  if (signals.some((x) => x.includes("stripe") || x.includes("payment"))) {
    out.push("Payment integrations should remain out of MVP-critical path unless monetization is a launch requirement.");
  }

  if (!integrations.length) {
    out.push("With minimal external dependencies, delivery risk is reduced and core platform reliability can be prioritized.");
  }

  return out;
}

function buildRiskMitigationPlan(riskInput, timelineWeeks, teamSize, coreFeatures, integrations) {
  const text = riskInput.toLowerCase();
  const featureSignals = toLowerList(coreFeatures);
  const integrationSignals = toLowerList(integrations);
  const rows = [];

  const keywordRows = [
    {
      keys: ["scope", "change request", "creep"],
      risk: "Scope creep across messaging and collaboration features",
      impact: "Schedule slippage and unstable release quality",
      mitigation: "Freeze MVP scope by week 2, enforce change-control gate, and move additions to Phase 2 backlog.",
    },
    {
      keys: ["security", "encryption", "privacy", "compliance", "standards"],
      risk: "Security/compliance gaps in authentication, data handling, or auditability",
      impact: "Production exposure, legal risk, and delayed launch approvals",
      mitigation: "Security checklist per sprint, mandatory threat modeling, and pre-release penetration testing.",
    },
    {
      keys: ["performance", "scale", "latency", "load"],
      risk: "Realtime performance degradation under concurrent usage",
      impact: "Message delays, disconnects, and poor user retention",
      mitigation: "Run load tests from week 5, optimize socket fan-out, and add autoscaling plus caching controls.",
    },
    {
      keys: ["integration", "third party", "twilio", "firebase"],
      risk: "External provider dependency failures or quota limits",
      impact: "OTP/notification interruptions and onboarding drop-offs",
      mitigation: "Use provider retry queues, fallback channels, and proactive quota/health monitoring.",
    },
    {
      keys: ["team", "resource", "bandwidth"],
      risk: "Insufficient engineering capacity for planned scope",
      impact: "Incomplete features or rushed QA near release",
      mitigation: "Prioritize critical path features and apply strict WIP limits with weekly capacity re-forecasting.",
    },
  ];

  for (const item of keywordRows) {
    if (item.keys.some((k) => text.includes(k))) {
      rows.push(item);
    }
  }

  if (timelineWeeks <= 8) {
    rows.push({
      risk: "Compressed delivery timeline",
      impact: "Reduced testing depth and higher regression probability",
      mitigation: "Reduce scope to strict MVP, shift non-critical integrations, and start QA automation by week 3.",
    });
  }

  const teamMatch = teamSize.match(/\d+/);
  if (teamMatch && Number(teamMatch[0]) <= 4) {
    rows.push({
      risk: "Small delivery team against chat platform complexity",
      impact: "Execution bottlenecks in parallel frontend/backend work",
      mitigation: "Use modular backlog slices and avoid parallel feature streams beyond team capacity.",
    });
  }

  if (featureSignals.some((x) => x.includes("media") || x.includes("file") || x.includes("attachment"))) {
    rows.push({
      risk: "Unsafe or oversized media uploads",
      impact: "Storage abuse, malware risk, and cost spikes",
      mitigation: "Apply MIME/size policies, virus scan workflow, and lifecycle cleanup jobs.",
    });
  }

  if (integrationSignals.some((x) => x.includes("twilio") || x.includes("firebase"))) {
    rows.push({
      risk: "Critical user flows dependent on third-party APIs",
      impact: "Login or notification failures during provider incidents",
      mitigation: "Introduce provider abstraction layer and circuit-breaker with graceful degradation paths.",
    });
  }

  if (!rows.length) {
    rows.push({
      risk: "Scope and delivery uncertainty",
      impact: "Potential release delay and quality inconsistencies",
      mitigation: "Define clear acceptance criteria, weekly demos, and risk review checkpoints.",
    });
  }

  return rows.map((r) => `Risk: ${r.risk} | Impact: ${r.impact} | Mitigation: ${r.mitigation}`);
}

function buildSecurityPlan(securityNeeds, integrations) {
  const text = securityNeeds.toLowerCase();
  const out = [
    "Enforce JWT access/refresh token rotation with secure session invalidation.",
    "Apply request validation, rate limiting, and abuse protection on all public endpoints.",
    "Use audit logging for auth events, admin actions, and moderation workflows.",
  ];

  if (text.includes("encryption") || text.includes("e2e")) {
    out.push("Introduce encryption strategy for data at rest and staged rollout path for end-to-end encryption.");
  }
  if (text.includes("gdpr") || text.includes("compliance") || text.includes("privacy")) {
    out.push("Add consent/retention controls, account deletion workflow, and data export capability.");
  }
  if (toLowerList(integrations).some((x) => x.includes("twilio") || x.includes("firebase"))) {
    out.push("Store third-party credentials in secret manager and rotate keys on a scheduled cadence.");
  }
  return out;
}

function buildDesignExecutionPlan(designDirection) {
  const text = designDirection.toLowerCase();
  const out = [
    "Establish a reusable design token system (color, spacing, typography, elevation) before feature screens.",
    "Ship mobile-first layouts first, then scale to desktop split-view experience.",
    "Define interaction states for message send, delivery, errors, retries, and empty states.",
  ];

  if (text.includes("accessibility") || text.includes("wcag")) {
    out.push("Include WCAG 2.2 AA checks for contrast, keyboard navigation, and screen-reader labeling.");
  }
  if (text.includes("premium") || text.includes("modern") || text.includes("whatsapp")) {
    out.push("Prioritize conversational readability: bubble spacing, avatar rhythm, and subtle motion for realtime feedback.");
  }
  return out;
}

function buildKpiPlan(successMetrics, timelineWeeks) {
  const text = successMetrics.toLowerCase();
  const out = [
    "Reliability KPI: crash-free sessions >= 99.5% during first month after launch.",
    "Latency KPI: p95 message delivery under 500 ms in normal traffic windows.",
    "Quality KPI: escaped-defect rate decreasing sprint-over-sprint from week 6 onward.",
  ];

  if (text.includes("retention")) {
    out.push("Adoption KPI: day-30 retention target tracked from launch cohort with weekly funnel review.");
  }
  if (text.includes("adoption") || text.includes("growth")) {
    out.push("Growth KPI: weekly active users and invite acceptance conversion monitored by channel.");
  }
  if (timelineWeeks <= 8) {
    out.push("Execution KPI: weekly scope-completion rate above 85% to protect schedule certainty.");
  }
  return out;
}

function buildWeeklyPlan(weeks, coreFeatures, advancedFeatures) {
  const weekTemplates = [
    { title: "Discovery and Requirement Lock", tasks: ["Finalize user journeys, acceptance criteria, and non-functional requirements.", "Define MVP boundaries for launch and a clear Phase 2 backlog.", "Prepare delivery board, risk register, and dependency map."] },
    { title: "Architecture and Project Setup", tasks: ["Bootstrap Node.js backend structure, frontend app shell, and repository standards.", "Define database schema v1 and socket communication contracts.", "Set up CI pipeline, linting, formatting, and baseline test harness."] },
    { title: "Authentication and Profile Foundation", tasks: ["Implement auth flows, session handling, and profile management APIs.", "Add client-side auth screens and protected route handling.", "Introduce role and permission primitives for future admin features."] },
    { title: "Conversation and Chat List Core", tasks: ["Build conversation creation/listing APIs for 1:1 and group-ready metadata.", "Implement chat list UI with unread state and sorting behavior.", "Add server-side validation, pagination, and error handling patterns."] },
    { title: "Realtime Messaging Engine", tasks: ["Implement message send/receive flow with Socket.IO and persistence consistency.", "Add message status events (sent/delivered/read) and idempotency checks.", "Run load baseline on websocket connections and message throughput."] },
    { title: "Groups, Presence, and Typing", tasks: ["Add participant management and group role controls.", "Implement online/offline presence, last-seen, and typing signals.", "Harden websocket reconnection and state resynchronization flows."] },
    { title: "Media and Search", tasks: ["Implement upload pipeline, attachment rendering, and file policy checks.", "Add conversation/message search with indexing strategy.", "Integrate object storage and background job processing for media tasks."] },
    { title: "Notifications and UX Polish", tasks: ["Integrate push notifications and preference-based delivery rules.", "Polish design system: spacing, typography, responsive breakpoints, and interaction feedback.", "Add accessibility improvements (labels, focus order, contrast adjustments)."] },
    { title: "QA Automation and Bug Stabilization", tasks: ["Expand unit and integration tests for critical message flows.", "Add end-to-end regression scenarios for auth, messaging, and groups.", "Resolve defects from QA cycle and tune error observability."] },
    { title: "Security and Performance Hardening", tasks: ["Implement rate limiting, abuse checks, and stricter payload validation.", "Run performance optimization for hot endpoints and socket fan-out paths.", "Complete security checklist and log/audit verification."] },
    { title: "UAT and Release Preparation", tasks: ["Run user acceptance tests and validate feature readiness against plan.", "Freeze MVP scope and prepare release notes plus rollback strategy.", "Train ops/support on runbooks and monitoring dashboards."] },
    { title: "Production Launch and Hypercare", tasks: ["Deploy to production with staged rollout and live telemetry watch.", "Monitor incident alerts, triage launch issues, and fix blockers quickly.", "Capture post-launch feedback and prioritize the first optimization sprint."] },
  ];

  const plan = [];
  const maxTemplates = weekTemplates.length;

  for (let i = 0; i < weeks; i += 1) {
    const base = i < maxTemplates ? weekTemplates[i] : {
      title: `Optimization Sprint ${i - maxTemplates + 1}`,
      tasks: [
        "Execute backlog features based on launch analytics and customer feedback.",
        "Improve reliability, observability, and deployment automation maturity.",
        "Prepare incremental releases with strict regression and performance checks.",
      ],
    };

    plan.push({ week: i + 1, title: base.title, tasks: [...base.tasks] });
  }

  const advanced = advancedFeatures.length ? advancedFeatures.join(", ") : "advanced enhancements";
  if (weeks > 8) {
    plan[weeks - 1].tasks.push(`Plan next cycle scope for ${advanced}.`);
  }

  if (coreFeatures.length && plan.length) {
    const targetIndex = plan.length >= 4 ? 3 : plan.length - 1;
    plan[targetIndex].tasks.push(`Priority implementation focus: ${coreFeatures.slice(0, 4).join(", ")}.`);
  }

  return plan;
}

function createPlanDocument(input) {
  const projectName = clean(input.projectName, 120) || "Untitled Project";
  const companyName = clean(input.companyName, 120) || "Not specified";
  const summary = clean(input.summary, 1000) || "No summary provided.";
  const targetUsers = clean(input.targetUsers, 1000) || "Not specified";
  const platforms = splitList(input.platforms);
  const timeline = clean(input.timeline, 120) || "Not specified";
  const budget = clean(input.budget, 120) || "Not specified";
  const teamSize = clean(input.teamSize, 120) || "Not specified";
  const coreFeatures = splitList(input.coreFeatures);
  const advancedFeatures = splitList(input.advancedFeatures);
  const integrations = splitList(input.integrations);
  const techStack = clean(input.techStack, 1500) || "Node.js + Express + modern frontend stack";
  const securityNeeds = clean(input.securityNeeds, 1500) || "Standard auth, validation, and secure storage";
  const designDirection = clean(input.designDirection, 1500) || "Clean modern product design with mobile-first UI";
  const successMetrics = clean(input.successMetrics, 1500) || "User adoption, retention, and response time";
  const risks = clean(input.risks, 1500) || "Scope growth and delivery timeline";
  const timelineWeeks = parseTimelineWeeks(timeline);
  const architectureRecommendations = buildArchitectureRecommendations(coreFeatures, integrations, securityNeeds, techStack);
  const modulePlan = buildModulePlan(coreFeatures);
  const dataModelPlan = buildDataModelPlan(coreFeatures);
  const apiPlan = buildApiPlan();
  const weeklyPlan = buildWeeklyPlan(timelineWeeks, coreFeatures, advancedFeatures);
  const executiveSummaryLines = buildExecutiveSummary(summary, targetUsers, coreFeatures, platforms);
  const audiencePlatformPlan = buildAudienceAndPlatformPlan(targetUsers, platforms);
  const deliveryConstraintPlan = buildDeliveryConstraintPlan(timelineWeeks, budget, teamSize);
  const featureScopePlan = buildFeatureScopePlan(coreFeatures, advancedFeatures);
  const integrationStrategy = buildIntegrationStrategy(integrations);
  const riskMitigationPlan = buildRiskMitigationPlan(risks, timelineWeeks, teamSize, coreFeatures, integrations);
  const securityPlan = buildSecurityPlan(securityNeeds, integrations);
  const designExecutionPlan = buildDesignExecutionPlan(designDirection);
  const kpiPlan = buildKpiPlan(successMetrics, timelineWeeks);

  const today = new Date().toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });

  return new Document({
    creator: "Project Plan Assistant",
    title: `${projectName} - Project Plan`,
    description: `Generated project plan for ${projectName}`,
    sections: [
      {
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 160 },
            children: [new TextRun({ text: "Project Plan", bold: true, size: 42 })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
            children: [new TextRun({ text: projectName, size: 32 })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 260 },
            children: [new TextRun({ text: `Prepared for: ${companyName} | Date: ${today}`, italics: true })],
          }),

          section("1. Executive Summary"),
          ...bullets(executiveSummaryLines),

          section("2. Project Context"),
          ...bullets(audiencePlatformPlan),
          ...bullets(deliveryConstraintPlan),

          section("3. Scope and Feature Plan"),
          sub("Core MVP Features"),
          ...bullets(featureScopePlan.outCore),
          sub("Phase 2 / Nice-to-Have Features"),
          ...bullets(featureScopePlan.outAdvanced),

          section("4. Recommended Node-Based Technical Architecture"),
          line("Technology direction is aligned to a Node.js-centered, modular, realtime-capable platform architecture."),
          sub("Architecture Blueprint"),
          ...bullets(architectureRecommendations),
          sub("Core Backend Modules"),
          ...bullets(modulePlan),
          sub("Initial Data Model"),
          ...bullets(dataModelPlan),
          sub("API and Socket Contract Scope"),
          ...bullets(apiPlan),
          sub("Integration Requirements"),
          ...bullets(integrationStrategy),

          section("5. Design and UX Direction"),
          ...bullets(designExecutionPlan),

          section("6. Security and Compliance"),
          ...bullets(securityPlan),

          section("7. Delivery Approach"),
          line("Week-by-week implementation roadmap:"),
          ...weeklyPlan.flatMap((item) => [
            sub(`Week ${item.week}: ${item.title}`),
            ...bullets(item.tasks),
          ]),

          section("8. Risks and Mitigation"),
          ...bullets(riskMitigationPlan),

          section("9. Success Metrics"),
          ...bullets(kpiPlan),

          section("10. Deliverables"),
          ...bullets([
            "Detailed project requirements",
            "UI/UX design system and high-fidelity screens",
            "Node.js codebase and API documentation",
            "Testing report and release checklist",
            "Deployment and handover documentation",
          ]),
        ],
      },
    ],
  });
}

app.post("/api/project-plan", async (req, res) => {
  try {
    const planDoc = createPlanDocument(req.body || {});
    const buffer = await Packer.toBuffer(planDoc);
    const safeName = clean(req.body?.projectName || "Project", 80)
      .replace(/[^a-zA-Z0-9-_ ]/g, "")
      .replace(/\s+/g, "-") || "Project";

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${safeName}-Project-Plan.docx"`);
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ error: "Failed to generate project plan." });
  }
});

app.listen(PORT, () => {
  // eslint-disable-next-line no-console
  console.log(`Project planner running on http://localhost:${PORT}`);
});
