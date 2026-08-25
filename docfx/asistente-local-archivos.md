# Acme Docs

> Product documentation for the Acme platform.

Other machine-readable surfaces for this site:
- [Full content](https://asistentelg.midominio.workers.dev/llms-full.txt)
- [Typed JSON index](https://asistentelg.midominio.workers.dev/index.json)

## Pages

- [Getting Started with Acme](https://asistentelg.midominio.workers.dev/getting-started.md) — This guide is for users who want to create their first workflow in Acme, a workflow automation platform.
- [Acme Pricing Plans](https://asistentelg.midominio.workers.dev/pricing.md) — Acme offers three pricing plans: Free, Pro, and Enterprise, each with varying levels of features and support.
- [Integrations](https://asistentelg.midominio.workers.dev/integrations.md) — Acme connects to various tools through built-in integrations and generic steps.
- [Security and Compliance](https://asistentelg.midominio.workers.dev/security.md) — Acme prioritizes security and compliance, offering SOC 2 Type II certification, GDPR compliance, and robust data encryption.
- [Acme FAQ](https://asistentelg.midominio.workers.dev/faq.md) — This page answers frequently asked questions about Acme, including self-hosting, API availability, and cancellation policies.
# Acme Docs

> Product documentation for the Acme platform.

# Getting Started with Acme

*Category: documentation*
*Topics: workflow automation, acme cli, yaml workflows, deployment, free account*

This guide is for users who want to create their first workflow on the Acme workflow automation platform. It walks through the process of installing the CLI, creating a workflow, and deploying it. The guide is designed to help users get started in under five minutes. Acme is suitable for users who need to automate workflows. The platform offers a free account option with limited workflow executions per month.

## Key points

- Install the Acme CLI with npm
- Create a workflow using a YAML file
- Deploy workflows with the Acme deploy command
- Free accounts have limited workflow executions
- Workflows can be triggered by schedule, webhook, or manual trigger

## Content

# Getting Started with Acme
Acme is a workflow automation platform. This guide walks you through creating your first workflow in under five minutes.
## Install the CLI
Run `npm install -g @acme/cli` and authenticate with `acme login`. The CLI stores a token in ~/.acme/config.
## Create a workflow
Workflows are YAML files. A minimal workflow has a trigger and one step:
* triggers: schedule, webhook, or manual
* steps: run a script, call an HTTP endpoint, or branch on a condition
Deploy with `acme deploy workflow.yaml`. Free accounts can run up to 1,000 workflow executions per month.

## Source

- Canonical URL: https://example.com/docs/getting-started
- Typed record: https://asistentelg.midominio.workers.dev/index.json


---

# Pricing Plans

*Category: pricing*
*Topics: pricing, plans, subscriptions, billing, support*

Acme offers three pricing plans: Free, Pro, and Enterprise, each with varying features and support levels. The plans are suitable for individuals and businesses with different needs. Annual billing is available with a 20% discount. The plans include a visual workflow editor and CLI.

## Key points

- Free plan: $0/month, 1,000 executions, community support
- Pro plan: $20/month, 50,000 executions, email support, unlimited projects
- Enterprise plan: custom pricing, SSO, SLA, dedicated support

## Content

# Pricing Plans
Acme has three plans.
## Free
* $0/month
* 1,000 executions
* Community support
* 1 project
## Pro
* $20/month
* 50,000 executions
* Email support
* Unlimited projects
* Audit logs
## Enterprise
* Custom pricing
* SSO
* SLA
* Dedicated support
* On-prem options
All plans include the visual workflow editor and the CLI. Annual billing saves 20%. There is no charge for failed executions.

## Source

- Canonical URL: https://example.com/pricing
- Typed record: https://asistentelg.midominio.workers.dev/index.json


---

# Integrations

*Category: documentation*
*Topics: integrations, automation, workflow, oauth, restapi*

Acme connects to various tools and services, allowing for seamless integration and automation. It offers built-in integrations with popular platforms and also supports custom connections through HTTP and Webhook triggers. This feature is designed for users who want to streamline their workflows and enhance productivity.

## Key points

- Built-in integrations with Slack, GitHub, Stripe, Salesforce, and Google Sheets
- Generic HTTP step for connecting to any REST API
- Webhook trigger for starting workflows from external events
- OAuth credentials are encrypted at rest

## Content

## Integrations
Acme connects to the tools you already use. Built-in integrations include Slack, GitHub, Stripe, Salesforce, and Google Sheets. Each integration is configured once under Settings → Integrations and can then be referenced from any workflow step. For anything without a built-in connector, use the generic HTTP step to call any REST API, or the Webhook trigger to start a workflow from an external event. OAuth credentials are encrypted at rest.

## Source

- Canonical URL: https://example.com/docs/integrations
- Typed record: https://asistentelg.midominio.workers.dev/index.json


---

# Security and Compliance

*Category: security*
*Topics: security, compliance, gdpr, encryption, penetration testing*

Acme prioritizes security and compliance, ensuring the protection of user data through various measures. The company is SOC 2 Type II certified and GDPR compliant, making it a reliable choice for enterprise customers. Acme's security features include data encryption and secure storage of secrets. The company also undergoes regular penetration testing to maintain the security of its systems.

## Key points

- SOC 2 Type II certified
- GDPR compliant
- Data encrypted in transit and at rest
- Secure storage of secrets in an isolated vault
- Annual third-party penetration testing

## Content

## Security & Compliance
Acme is SOC 2 Type II certified and GDPR compliant. All data is encrypted in transit (TLS 1.3) and at rest (AES-256). Secrets used in workflows are stored in an isolated vault and never logged. Enterprise customers can enable SAML single sign-on, scoped API tokens, and IP allowlisting. We undergo annual third-party penetration testing and publish a status page at status.example.com.

## Source

- Canonical URL: https://example.com/security
- Typed record: https://asistentelg.midominio.workers.dev/index.json


---

# Acme FAQ

*Category: documentation*
*Topics: acme, faq, selfhosting, api, billing, cancellation*

This page provides answers to frequently asked questions about Acme, including self-hosting, API availability, and cancellation policies. It is intended for users of Acme's services. The FAQ covers various topics, including deployment options and billing. Acme offers flexible solutions for its users.

## Key points

- Acme offers self-hosting through Enterprise plans
- A REST API and CLI are available for dashboard access
- Workflow scripts support JavaScript and Python
- Cancellation is available through the Settings → Billing menu

## Content

# FAQ
## Can I self-host Acme?
Yes — Enterprise plans include an on-prem deployment option packaged as a container.
## Do you have an API?
Yes, everything in the dashboard is available through the REST API and the CLI.
## What languages do workflow scripts support?
JavaScript and Python today.
## How do I cancel?
From Settings → Billing. Cancellation takes effect at the end of the current billing period; we don't offer prorated refunds.

## Source

- Canonical URL: https://example.com/faq
- Typed record: https://asistentelg.midominio.workers.dev/index.json


---
{
  "protocol": "agent-visibility/0.1",
  "site": {
    "name": "Acme Docs",
    "description": "Product documentation for the Acme platform."
  },
  "generatedAt": "2026-08-25T07:17:41.859Z",
  "surfaces": {
    "llmsTxt": "https://asistentelg.midominio.workers.dev/llms.txt",
    "llmsFullTxt": "https://asistentelg.midominio.workers.dev/llms-full.txt",
    "json": "https://asistentelg.midominio.workers.dev/index.json",
    "pageMarkdown": "https://asistentelg.midominio.workers.dev/{slug}.md",
    "robots": "https://asistentelg.midominio.workers.dev/robots.txt"
  },
  "pages": [
    {
      "slug": "getting-started",
      "url": "https://example.com/docs/getting-started",
      "title": "Getting Started with Acme",
      "summary": "This guide is for users who want to create their first workflow in Acme, a workflow automation platform. It walks through the process of installing the CLI, creating a workflow, and deploying it. The guide is designed to be completed in under five minutes. Acme is suitable for users who need to automate workflows. Free accounts have limitations on workflow executions.",
      "keyPoints": [
        "Install the Acme CLI with npm",
        "Create a minimal workflow with a trigger and one step",
        "Deploy workflows with the Acme CLI",
        "Free accounts have 1,000 workflow executions per month"
      ],
      "topics": [
        "acme",
        "workflow automation",
        "cli",
        "yaml"
      ],
      "category": "documentation",
      "updatedAt": "2026-08-25T07:17:22.857Z",
      "sources": {
        "markdown": "https://asistentelg.midominio.workers.dev/getting-started.md",
        "canonical": "https://example.com/docs/getting-started"
      }
    },
    {
      "slug": "pricing",
      "url": "https://example.com/pricing",
      "title": "Acme Pricing Plans",
      "summary": "Acme offers three pricing plans: Free, Pro, and Enterprise, catering to different user needs and providing various features such as executions, support, and projects. The plans are designed for individuals, professionals, and large organizations. Annual billing is available with a 20% discount.",
      "keyPoints": [
        "Free plan: $0/month, 1,000 executions, community support",
        "Pro plan: $20/month, 50,000 executions, email support, unlimited projects",
        "Enterprise plan: custom pricing, SSO, SLA, dedicated support"
      ],
      "topics": [
        "pricing",
        "plans",
        "executions",
        "support",
        "billing"
      ],
      "category": "pricing",
      "updatedAt": "2026-08-25T07:17:27.412Z",
      "sources": {
        "markdown": "https://asistentelg.midominio.workers.dev/pricing.md",
        "canonical": "https://example.com/pricing"
      }
    },
    {
      "slug": "integrations",
      "url": "https://example.com/docs/integrations",
      "title": "Integrations",
      "summary": "Acme connects to various tools and services, allowing for streamlined workflows and automated processes. This page is for users looking to integrate their existing tools with Acme. The platform offers built-in integrations with popular services and also provides options for custom integrations.",
      "keyPoints": [
        "Built-in integrations with Slack, GitHub, Stripe, Salesforce, and Google Sheets",
        "Generic HTTP step for connecting to any REST API",
        "Webhook trigger for starting workflows from external events",
        "OAuth credentials are encrypted at rest"
      ],
      "topics": [
        "integrations",
        "api",
        "workflow",
        "automation",
        "security"
      ],
      "category": "documentation",
      "updatedAt": "2026-08-25T07:17:32.515Z",
      "sources": {
        "markdown": "https://asistentelg.midominio.workers.dev/integrations.md",
        "canonical": "https://example.com/docs/integrations"
      }
    },
    {
      "slug": "security",
      "url": "https://example.com/security",
      "title": "Security and Compliance",
      "summary": "Acme prioritizes security and compliance, ensuring the protection of customer data through various measures. The company is SOC 2 Type II certified and GDPR compliant, making it a reliable choice for enterprise customers. Acme's security features include encryption, isolated vaults for secrets, and annual penetration testing. This page provides an overview of Acme's security and compliance features, suitable for enterprise customers and individuals concerned about data protection.",
      "keyPoints": [
        "SOC 2 Type II certified",
        "GDPR compliant",
        "Data encrypted in transit and at rest",
        "Secrets stored in an isolated vault",
        "Annual third-party penetration testing"
      ],
      "topics": [
        "security",
        "compliance",
        "gdpr",
        "encryption",
        "penetration testing"
      ],
      "category": "security",
      "updatedAt": "2026-08-25T07:17:36.491Z",
      "sources": {
        "markdown": "https://asistentelg.midominio.workers.dev/security.md",
        "canonical": "https://example.com/security"
      }
    },
    {
      "slug": "faq",
      "url": "https://example.com/faq",
      "title": "Acme FAQ",
      "summary": "This page answers frequently asked questions about Acme, including self-hosting, API availability, supported languages, and cancellation procedures. It is intended for users of Acme who have questions about its features and usage. The page provides concise answers to common questions, helping users understand how to use Acme effectively. Acme offers various options for its users, including on-prem deployment and API access.",
      "keyPoints": [
        "Acme offers self-hosting through Enterprise plans",
        "Acme has a REST API and CLI",
        "Workflow scripts support JavaScript and Python",
        "Cancellation is done through Settings → Billing"
      ],
      "topics": [
        "acme",
        "faq",
        "selfhosting",
        "api",
        "workflow",
        "cancellation"
      ],
      "category": "documentation",
      "updatedAt": "2026-08-25T07:17:41.859Z",
      "sources": {
        "markdown": "https://asistentelg.midominio.workers.dev/faq.md",
        "canonical": "https://example.com/faq"
      }
    }
  ]
}

# Robots directives for AI agents and crawlers.
# This site intentionally welcomes AI agents — see /llms.txt.

User-agent: GPTBot
Allow: /
Content-Signal: ai-input=yes, search=yes, ai-train=no

User-agent: OAI-SearchBot
Allow: /
Content-Signal: ai-input=yes, search=yes, ai-train=no

User-agent: ChatGPT-User
Allow: /
Content-Signal: ai-input=yes, search=yes, ai-train=no

User-agent: ClaudeBot
Allow: /
Content-Signal: ai-input=yes, search=yes, ai-train=no

User-agent: Claude-User
Allow: /
Content-Signal: ai-input=yes, search=yes, ai-train=no

User-agent: PerplexityBot
Allow: /
Content-Signal: ai-input=yes, search=yes, ai-train=no

User-agent: Google-Extended
Allow: /
Content-Signal: ai-input=yes, search=yes, ai-train=no

User-agent: Applebot-Extended
Allow: /
Content-Signal: ai-input=yes, search=yes, ai-train=no

User-agent: Bytespider
Allow: /
Content-Signal: ai-input=yes, search=yes, ai-train=no

User-agent: CCBot
Allow: /
Content-Signal: ai-input=yes, search=yes, ai-train=no

User-agent: *
Allow: /
Content-Signal: ai-input=yes, search=yes, ai-train=no

# Machine-readable indexes for agents:
# - https://asistentelg.midominio.workers.dev/llms.txt
# -https://asistentelg.midominio.workers.dev/index.json

{
  "@context": "https://schema.org",
  "@type": "WebSite",
  "name": "Acme Docs",
  "description": "Product documentation for the Acme platform.",
  "url": "https://asistentelg.midominio.workers.dev",
  "mainEntity": {
    "@type": "ItemList",
    "itemListElement": [
      {
        "@type": "ListItem",
        "position": 1,
        "url": "https://asistentelg.midominio.workers.dev/getting-started.md",
        "name": "Getting Started with Acme"
      },
      {
        "@type": "ListItem",
        "position": 2,
        "url": "https://asistentelg.midominio.workers.dev/pricing.md",
        "name": "Pricing Plans"
      },
      {
        "@type": "ListItem",
        "position": 3,
        "url": "https://asistentelg.midominio.workers.dev/integrations.md",
        "name": "Integrations"
      },
      {
        "@type": "ListItem",
        "position": 4,
        "url": "https://asistentelg.midominio.workers.dev/security.md",
        "name": "Security and Compliance"
      },
      {
        "@type": "ListItem",
        "position": 5,
        "url": "https://asistentelg.midominio.workers.dev/faq.md",
        "name": "Acme FAQ"
      }
    ]
  }
}













