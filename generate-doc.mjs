import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, Table, TableRow, TableCell,
  WidthType, ShadingType, PageBreak, UnderlineType
} from "docx";
import { writeFileSync } from "fs";

const GOLD = "B8860B";
const NAVY = "0E1A35";
const DARK = "1A1A2E";
const GRAY = "6B7280";
const LINE = "E5E7EB";
const WHITE = "FFFFFF";

function pageTitle(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 52, color: NAVY, font: "Calibri" })],
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD, space: 6 } }
  });
}

function sectionLabel(text) {
  return new Paragraph({
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 18, color: GOLD, font: "Calibri", characterSpacing: 100 })],
    spacing: { before: 360, after: 60 }
  });
}

function sectionHeading(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 34, color: NAVY, font: "Calibri" })],
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 120, after: 100 }
  });
}

function subHeading(num, text) {
  return new Paragraph({
    children: [
      new TextRun({ text: `${num}  `, bold: true, size: 22, color: GOLD, font: "Calibri" }),
      new TextRun({ text, bold: true, size: 22, color: NAVY, font: "Calibri" })
    ],
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 240, after: 80 }
  });
}

function body(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, color: DARK, font: "Calibri" })],
    spacing: { before: 60, after: 60 },
    indent: { left: 0 }
  });
}

function lead(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 24, color: NAVY, font: "Calibri" })],
    spacing: { before: 80, after: 80 }
  });
}

function bullet(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 21, color: DARK, font: "Calibri" })],
    bullet: { level: 0 },
    spacing: { before: 30, after: 30 }
  });
}

function subBullet(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 20, color: GRAY, font: "Calibri" })],
    bullet: { level: 1 },
    spacing: { before: 20, after: 20 }
  });
}

function divider() {
  return new Paragraph({
    children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } },
    spacing: { before: 200, after: 200 }
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function statRow(items) {
  return new Paragraph({
    children: items.flatMap(([val, label]) => [
      new TextRun({ text: val, bold: true, size: 28, color: GOLD, font: "Calibri" }),
      new TextRun({ text: `  ${label}     `, size: 20, color: GRAY, font: "Calibri" })
    ]),
    spacing: { before: 120, after: 120 }
  });
}

function qaItem(q, a) {
  return [
    new Paragraph({
      children: [new TextRun({ text: `Q: ${q}`, bold: true, size: 22, color: NAVY, font: "Calibri" })],
      spacing: { before: 200, after: 60 }
    }),
    new Paragraph({
      children: [new TextRun({ text: `A: ${a}`, size: 21, color: DARK, font: "Calibri" })],
      spacing: { before: 0, after: 100 }
    })
  ];
}

// ─────────────────────────────────────────────────────────
// BUILD DOCUMENT
// ─────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 22, color: DARK } }
    }
  },
  sections: [{
    properties: {
      page: { margin: { top: 1080, bottom: 1080, left: 1200, right: 1200 } }
    },
    children: [

      // ═══════════════════════════════════════════════
      // COVER
      // ═══════════════════════════════════════════════
      new Paragraph({
        children: [new TextRun({ text: "BrillBrains Consultants Pvt. Ltd.", bold: true, size: 72, color: NAVY, font: "Calibri" })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 800, after: 200 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "Strategic & Creative Forever", size: 36, color: GOLD, font: "Calibri", italics: true })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 200 }
      }),
      new Paragraph({
        children: [],
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 0 } },
        spacing: { before: 200, after: 200 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "Complete Website Content Document", size: 28, color: GRAY, font: "Calibri" })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 100 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "All pages | All sections | All copy", size: 24, color: GRAY, font: "Calibri" })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 800 }
      }),
      pageBreak(),

      // ═══════════════════════════════════════════════
      // NAVIGATION BAR
      // ═══════════════════════════════════════════════
      pageTitle("NAVIGATION BAR (All Pages)"),
      sectionLabel("Brand Logo Area"),
      body("Logo: BrillBrains logo (bb-logo.png)"),
      body("Brand Name: BrillBrains"),
      body("Tagline: Strategic & Creative Forever"),
      sectionLabel("Primary Navigation Links"),
      bullet("Home"),
      bullet("About Us"),
      bullet("Our Clients"),
      bullet("Services  (dropdown)"),
      subBullet("Brand Consulting"),
      subBullet("Business Management"),
      subBullet("Tech Consulting"),
      subBullet("Training Programs"),
      bullet("Careers"),
      bullet("Connect  [CTA button]"),
      divider(),

      // ═══════════════════════════════════════════════
      // PAGE 1 — HOME
      // ═══════════════════════════════════════════════
      pageBreak(),
      pageTitle("PAGE 1 — HOME (index.html)"),

      sectionLabel("Hero Section"),
      sectionHeading("Strategic & Creative Forever"),
      body("Eyebrow: Since 2011 - Jewellery Industry"),
      body("Heading: Strategic & Creative Forever"),
      body("Body: BrillBrains is a Business Growth Consultant firm dedicated to the Jewellery Industry - strategically transforming brands and businesses through brand consulting, management consulting, tech consulting, and professional training."),
      body("Button 1: Explore Services"),
      body("Button 2: Talk to us"),
      sectionLabel("Stats"),
      statRow([["15+", "Years of Business"], ["4", "Core Practices"], ["10+", "Brands Empowered"]]),

      sectionLabel("Marquee / Ticker Strip"),
      body("Brand Consulting  •  Management Consulting  •  Tech Consulting  •  Professional Training  (repeating)"),

      sectionLabel("Our Clients Section"),
      sectionHeading("Trusted by Jewellery businesses worldwide"),
      body("Eyebrow: Our Clients"),
      body("Body: We partner with retail jewellers, wholesalers, and SMEs across the globe to build stronger brands, cleaner operations, and empowered teams."),
      body("Client logos: 4 placeholder slots (to be replaced with actual client logos)"),
      body("Button: See More  →  links to clients.html"),

      sectionLabel("About Split Section"),
      sectionHeading("Business growth consulting for Jewellery businesses"),
      body("Eyebrow: About"),
      lead("We serve retail jewellers and wholesalers worldwide as management and brand consultants, with deep specialization in operational documentation, systems, processes, technology, and people development."),
      body("BrillBrains also has a division for customized and professional business training programs to boost organizational efficiency and productivity."),
      statRow([["2011", "Established"], ["SMEs", "Focus"], ["Global", "Retail & Wholesale"]]),

      sectionLabel("What We Do — Service Pillars"),
      sectionHeading("The service ecosystem"),
      body("Eyebrow: What We Do"),
      body("Sub-heading: Four focused practices designed to elevate a Jewellery business end to end."),
      subHeading("01", "Brand Consulting"),
      bullet("Brand consultancy"), bullet("Creative services"), bullet("Content"), bullet("Digital marketing"), bullet("Marketing strategy"), bullet("Video marketing"), bullet("Consultancy"),
      subHeading("02", "Business Management Consulting"),
      bullet("Business management mastery"), bullet("Process documentation"), bullet("Streamlining"), bullet("Systems"), bullet("Work profiles"), bullet("Tracking"), bullet("Empowered teams"), bullet("Autopilot business"), bullet("Expansion strategy"),
      subHeading("03", "Tech Consulting"),
      bullet("Digital transformation"), bullet("Cloud solutions"), bullet("CRM and ERP setup"), bullet("Automation"), bullet("Analytics"), bullet("Cybersecurity"), bullet("IT infrastructure"), bullet("Custom software"), bullet("The Web Spiders"),
      subHeading("04", "Training Programs"),
      bullet("Management training"), bullet("Jewellery Graduate"), bullet("Culture transformation"), bullet("Team building"), bullet("Marketing mastery"), bullet("Digital Shakti"), bullet("Leadership"), bullet("Tours & exhibitions"),

      sectionLabel("Vision Section"),
      sectionHeading("Build a Powerful Team — Build an Empire — Win Bigger"),
      body("Eyebrow: Our Vision"),
      body("To be the trusted Business Growth Consultant firm in the Jewellery Industry. We are committed to making business owners, entrepreneurs, leaders, and team members unstoppable by strategically helping, guiding, consulting, and coaching them to achieve personal and professional goals."),

      sectionLabel("Founder Section"),
      sectionHeading("The brain behind Brill Brains"),
      body("Eyebrow: Founder"),
      body("Quote: BrillBrains is shaped by a founder's vision for structured Jewellery business growth: stronger brands, documented operations, empowered teams, and scalable systems."),
      body("Body: The founder section anchors the team around clarity, implementation, and long-term partnership. Replace this portrait and caption with the founder's real photo and name whenever ready."),
      bullet("Business Growth Strategy"),
      bullet("Jewellery Industry Focus"),
      bullet("Systems & Team Building"),

      sectionLabel("Strategic Partnership Section"),
      sectionHeading("Power to Performance"),
      body("Eyebrow: Strategic Partnership"),
      body("Body: BrillBrains and DG Trainings partner to deliver world-class professional training programs for Jewellery businesses — combining consulting expertise with proven development frameworks."),
      body("Partner 1: BrillBrains  —  Business Growth Consulting"),
      body("Connector: ×"),
      body("Partner 2: DG Trainings  —  Professional Training Programs"),

      sectionLabel("How We Work — Process Steps"),
      sectionHeading("From audit to empire"),
      body("Eyebrow: How We Work"),
      subHeading("01", "Discover"),
      body("Deep brand audit - operations, identity, market position, and team mindset."),
      subHeading("02", "Design"),
      body("Strategy, systems, and creative direction tailored to your Jewellery business."),
      subHeading("03", "Deploy"),
      body("Documentation, processes, and training rolled out across the organization."),
      subHeading("04", "Develop"),
      body("Long-term partnership - continuous improvement, reporting, and scaling."),

      sectionLabel("Contact CTA Band"),
      sectionHeading("Ready to build your empire?"),
      body("Eyebrow: Get in Touch"),
      body("Body: Reach out to start a conversation about your Jewellery brand, operations, and growth goals. We respond within one business day."),
      body("Button 1: Connect Now"),
      body("Button 2: +91 73041 92092"),
      divider(),

      // ═══════════════════════════════════════════════
      // PAGE 2 — ABOUT US
      // ═══════════════════════════════════════════════
      pageBreak(),
      pageTitle("PAGE 2 — ABOUT US (about.html)"),

      sectionLabel("Hero Section"),
      sectionHeading("About BrillBrains Consultants Pvt Ltd"),
      body("Eyebrow: About Us"),
      body("Body: Welcome to BrillBrains - strategically transforming brands and businesses since 2011."),

      sectionLabel("Who We Are Split Section"),
      sectionHeading("Focused on the Jewellery industry"),
      body("Eyebrow: Who We Are"),
      lead("BrillBrains is a Business Growth Consultant firm dedicated to the Jewellery Industry."),
      body("We serve retail jewellers and wholesalers worldwide as management and brand consultants, explicitly specializing in operational documentation, systems, and processes to help SMEs achieve operational excellence."),
      body("Our training division creates customized and professional business training programs to boost the efficiency and productivity of the overall organization."),

      sectionLabel("Four Pillars Cards"),
      subHeading("01", "Business Growth"),
      body("Structured consulting that connects brand, business operations, technology, and people."),
      subHeading("02", "Operational Excellence"),
      body("Documentation, processes, work profiles, reporting, and repeatable systems for SMEs."),
      subHeading("03", "Team Empowerment"),
      body("Training and leadership development that helps owners, leaders, and teams perform with clarity."),
      subHeading("04", "Long-term Partnership"),
      body("Guidance designed around implementation, accountability, and sustainable growth."),

      sectionLabel("Message / Vision Card"),
      sectionHeading("Simple Yet powerful"),
      body("Eyebrow: Message"),
      body("Build a Powerful Team. Build an Empire. Be Unstoppable."),
      divider(),

      // ═══════════════════════════════════════════════
      // PAGE 3 — OUR CLIENTS
      // ═══════════════════════════════════════════════
      pageBreak(),
      pageTitle("PAGE 3 — OUR CLIENTS (clients.html)"),

      sectionLabel("Hero Section"),
      sectionHeading("Partners in Jewellery business growth"),
      body("Eyebrow: Our Clients"),
      body("Body: BrillBrains works with Jewellery retailers, wholesalers, SMEs, entrepreneurs, leaders, and teams that want structured growth, stronger brands, cleaner operations, and measurable performance."),

      sectionLabel("Client Profile Split Section"),
      sectionHeading("Built for ambitious Jewellery businesses"),
      body("Eyebrow: Client Profile"),
      lead("Our client work is centered on retail jewellers and wholesalers worldwide."),
      body("We support businesses that want to move from owner-dependent execution to documented, process-driven, team-powered operations. Engagements often include brand repositioning, process documentation, department setup, performance tracking, technology adoption, and training."),

      sectionLabel("Client Logo Carousel"),
      body("12 client brand logo slots (scrolling carousel — to be replaced with actual client logos)"),

      sectionLabel("Client Type Cards"),
      subHeading("01", "Retail Jewellers"),
      body("Customer-facing Jewellery businesses building stronger brands, better teams, and consistent sales experiences."),
      subHeading("02", "Wholesalers"),
      body("Jewellery businesses that need sharper systems, documentation, expansion strategy, and B2B excellence."),
      subHeading("03", "SME Owners"),
      body("Entrepreneurs looking to reduce day-to-day dependency and create autopilot business systems."),
      subHeading("04", "Leadership Teams"),
      body("Managers and teams that need clarity in roles, reporting, communication, and performance ownership."),

      sectionLabel("Outcomes Section"),
      sectionHeading("What clients come to build"),
      body("Eyebrow: Outcomes"),
      body("Luxury-ready brand identity, documented operations, empowered teams, technology-enabled workflows, and expansion-ready business systems."),
      divider(),

      // ═══════════════════════════════════════════════
      // PAGE 4 — SERVICES
      // ═══════════════════════════════════════════════
      pageBreak(),
      pageTitle("PAGE 4 — SERVICES (services.html)"),

      sectionLabel("Hero Section"),
      sectionHeading("Four consulting practices — One growth system"),
      body("Eyebrow: Services"),
      body("Body: Every engagement at BrillBrains is designed to elevate the Jewellery business end to end - from identity to operations, technology, and people."),

      sectionLabel("Service Pillar Grid"),
      subHeading("01", "Brand Consulting"),
      bullet("Brand consultancy"), bullet("Creative services"), bullet("Content"), bullet("Digital marketing"), bullet("Marketing strategy"), bullet("Video marketing"), bullet("Consultancy"),
      subHeading("02", "Business Management Consulting"),
      bullet("Business management mastery"), bullet("Process documentation"), bullet("Streamlining"), bullet("Systems"), bullet("Work profiles"), bullet("Tracking"), bullet("Empowered teams"), bullet("Autopilot business"), bullet("Expansion strategy"),
      subHeading("03", "Tech Consulting"),
      bullet("Digital transformation"), bullet("Cloud solutions"), bullet("CRM and ERP setup"), bullet("Automation"), bullet("Analytics"), bullet("Cybersecurity"), bullet("IT infrastructure"), bullet("Custom software"), bullet("The Web Spiders"),
      subHeading("04", "Training Programs"),
      bullet("Management training"), bullet("Jewellery Graduate"), bullet("Culture transformation"), bullet("Team building"), bullet("Marketing mastery"), bullet("Orientation"), bullet("Digital Shakti"), bullet("Leadership"), bullet("Tours"), bullet("Exhibitions"),
      divider(),

      // ═══════════════════════════════════════════════
      // PAGE 5 — BRAND CONSULTING
      // ═══════════════════════════════════════════════
      pageBreak(),
      pageTitle("PAGE 5 — BRAND CONSULTING (brand-consulting.html)"),

      sectionLabel("Hero Section"),
      sectionHeading("Brand consulting for Jewellery businesses"),
      body("Eyebrow: Services / Brand Consulting"),
      body("Body: Crafting a unique brand identity that elevates your Jewellery business to a trusted, luxury standard - from audit and strategy to identity, content, marketing, video, and web."),

      sectionLabel("Service Detail Cards"),
      subHeading("01", "Brand Consultancy"),
      bullet("Brand audit"), bullet("Brand strategy"), bullet("Brand promise"), bullet("Vision and mission"), bullet("Brand voice"), bullet("Positioning"), bullet("Identity"), bullet("Brand guidelines"),
      subHeading("02", "Creative Services"),
      bullet("Logo"), bullet("Social media creatives"), bullet("Festival wishes"), bullet("Outdoor advertising"), bullet("Marketing collaterals"), bullet("Event creatives"), bullet("Print media"), bullet("Web banners"),
      subHeading("03", "Content Creators"),
      bullet("Campaigns"), bullet("Ready templates"), bullet("Festival copies"), bullet("Videos"), bullet("Reels"), bullet("Presentations"), bullet("Speeches"), bullet("Newsletters"), bullet("Product descriptions"), bullet("Website content"),
      subHeading("04", "Digital Marketing"),
      bullet("Social media management"), bullet("SEO"), bullet("SEM"), bullet("PPC"), bullet("WhatsApp marketing"), bullet("Paid media"), bullet("Google My Business"), bullet("Content marketing"), bullet("Google Ads"),
      subHeading("05", "Marketing Strategy"),
      bullet("Competitor analysis"), bullet("Campaign planning"), bullet("Client engagement"), bullet("Engagement activities"), bullet("Events and exhibition planning"), bullet("Promotional schemes"), bullet("Marketing tools"),
      subHeading("06", "Video Marketing"),
      bullet("Promotional videos"), bullet("Reels"), bullet("Short videos"), bullet("Corporate videos"), bullet("Character animated videos"), bullet("Tutorial videos"),
      subHeading("07", "Consultancy"),
      bullet("Photoshoot and video shoot guidance"), bullet("Visual merchandising"), bullet("Brand and marketing department setup"), bullet("Gifts"), bullet("Packaging"), bullet("Media buying"),
      divider(),

      // ═══════════════════════════════════════════════
      // PAGE 6 — BUSINESS MANAGEMENT CONSULTING
      // ═══════════════════════════════════════════════
      pageBreak(),
      pageTitle("PAGE 6 — BUSINESS MANAGEMENT CONSULTING (business-management-consulting.html)"),

      sectionLabel("Hero Section"),
      sectionHeading("Business management for scalable operations"),
      body("Eyebrow: Services / Business Management Consulting"),
      body("Body: Optimizing operations, systems, and strategies to drive growth and efficiency in the Jewellery industry through transparent reporting, holistic frameworks, and adaptable systems."),

      sectionLabel("Service Detail Cards"),
      subHeading("01", "Business Management Mastery"),
      bullet("Owner clarity"), bullet("Team accountability"), bullet("Business rhythm"), bullet("Growth decision-making"),
      subHeading("02", "Process Documentation"),
      bullet("SOPs"), bullet("Manuals"), bullet("Workflows"), bullet("Checklists"), bullet("Repeatable documentation"),
      subHeading("03", "Existing Process Streamlining"),
      bullet("Current workflow audits"), bullet("Friction removal"), bullet("Duplication reduction"), bullet("Clear ownership"),
      subHeading("04", "System Designing"),
      bullet("Daily execution systems"), bullet("Departmental coordination"), bullet("Measurable outcomes"),
      subHeading("05", "Work Profiles"),
      bullet("Clear roles"), bullet("Responsibilities"), bullet("Reporting lines"), bullet("Team expectations"),
      subHeading("06", "Tracking"),
      bullet("KPIs"), bullet("Dashboards"), bullet("Review rhythms"), bullet("Continuous-improvement loops"),
      subHeading("07", "Empowered Teams"),
      bullet("Communication"), bullet("Ownership culture"), bullet("Training"), bullet("Leadership alignment"),
      subHeading("08", "Autopilot Business"),
      bullet("Owner dependency reduction"), bullet("Process consistency"), bullet("Structured operations"),
      subHeading("09", "Expansion Strategy"),
      bullet("Scalable systems"), bullet("Resource planning"), bullet("Operating models"), bullet("New market growth"),
      divider(),

      // ═══════════════════════════════════════════════
      // PAGE 7 — TECH CONSULTING
      // ═══════════════════════════════════════════════
      pageBreak(),
      pageTitle("PAGE 7 — TECH CONSULTING (tech-consulting.html)"),

      sectionLabel("Hero Section"),
      sectionHeading("Technology that strengthens the business system"),
      body("Eyebrow: Services / Tech Consulting"),
      body("Body: Tech consulting for Jewellery businesses that want cleaner data, smarter workflows, better customer management, safer infrastructure, and custom digital tools."),

      sectionLabel("Service Detail Cards"),
      subHeading("01", "Digital Transformation"),
      bullet("Manual workflow modernization"), bullet("Connected teams"), bullet("Practical digital systems"),
      subHeading("02", "Cloud Solutions"),
      bullet("Cloud-based tools"), bullet("Storage"), bullet("Collaboration"), bullet("Access control"), bullet("Scalable infrastructure planning"),
      subHeading("03", "CRM & ERP Setup"),
      bullet("Customer management"), bullet("Inventory"), bullet("Operations"), bullet("Reporting"), bullet("Department workflows"),
      subHeading("04", "Automation Tools"),
      bullet("Automated reminders"), bullet("Approvals"), bullet("Reporting"), bullet("Lead flows"), bullet("Repetitive task reduction"),
      subHeading("05", "Data & Analytics"),
      bullet("Dashboards"), bullet("Business intelligence"), bullet("Sales tracking"), bullet("Performance insights"), bullet("Decision support"),
      subHeading("06", "Cybersecurity"),
      bullet("Security practices"), bullet("Access policies"), bullet("Backups"), bullet("Awareness"), bullet("Safer digital operations"),
      subHeading("07", "IT Infrastructure"),
      bullet("Network"), bullet("Devices"), bullet("Software stack"), bullet("Support planning"), bullet("Reliable technology foundations"),
      subHeading("08", "Custom Software"),
      bullet("Tailored web solutions"), bullet("Internal tools"), bullet("Integrations"), bullet("Applications for your operating model"),
      subHeading("09", "The Web Spiders"),
      bullet("Domain name"), bullet("Web hosting"), bullet("Website designing"), bullet("Website development"),
      divider(),

      // ═══════════════════════════════════════════════
      // PAGE 8 — TRAINING PROGRAMS
      // ═══════════════════════════════════════════════
      pageBreak(),
      pageTitle("PAGE 8 — TRAINING PROGRAMS (training-programs.html)"),

      sectionLabel("Hero Section"),
      sectionHeading("Professional training for Jewellery teams"),
      body("Eyebrow: Services / Training Programs"),
      body("Body: Customized training programs designed to improve productivity, leadership, culture, sales confidence, digital skills, and execution across the organization."),

      sectionLabel("Training Program Cards"),
      subHeading("01", "Management Training"),
      bullet("Leadership capability"), bullet("Manager development"), bullet("Key team member training"),
      subHeading("02", "The Jewel Graduate"),
      bullet("Foundation building"), bullet("Jewellery business basics"), bullet("Professional growth skills"),
      subHeading("03", "Life & Work Culture Transformation"),
      bullet("Mindset"), bullet("Communication"), bullet("Discipline"), bullet("Ownership"), bullet("Workplace culture development"),
      subHeading("04", "Team Building"),
      bullet("Collaboration"), bullet("Trust"), bullet("Shared responsibility"), bullet("Team alignment"),
      subHeading("05", "Marketing Mastery"),
      bullet("Marketing clarity"), bullet("Campaigns"), bullet("Client engagement"), bullet("Jewellery business promotion"),
      subHeading("06", "Outstanding Orientation"),
      bullet("Onboarding systems"), bullet("Orientation structure"), bullet("Faster team productivity"),
      subHeading("07", "Digital Shakti"),
      bullet("Digital confidence"), bullet("Smartphone training"), bullet("Practical technology skills"),
      subHeading("08", "Leadership Training"),
      bullet("Decision-making"), bullet("Accountability"), bullet("Communication"), bullet("Motivation"), bullet("Team leadership"),
      subHeading("09", "Business Tour Experts"),
      bullet("B2B sales tours"), bullet("Travel discipline"), bullet("Buyer meetings"), bullet("Field execution"),
      subHeading("10", "The Exhibition Excellence"),
      bullet("Pre-exhibition preparation"), bullet("Booth discipline"), bullet("Lead handling"), bullet("Follow-up"), bullet("Conversion focus"),
      divider(),

      // ═══════════════════════════════════════════════
      // PAGE 9 — CAREERS
      // ═══════════════════════════════════════════════
      pageBreak(),
      pageTitle("PAGE 9 — CAREERS (careers.html)"),

      sectionLabel("Hero Section"),
      sectionHeading("Grow with BrillBrains"),
      body("Eyebrow: Careers"),
      body("Body: We're looking for driven, creative, and strategic minds to join our team. Help us build stronger brands, smarter systems, and empowered teams for Jewellery businesses worldwide."),

      sectionLabel("Why BrillBrains Split Section"),
      sectionHeading("A team built for impact"),
      body("Eyebrow: Why BrillBrains"),
      lead("At BrillBrains, every role contributes directly to transforming real businesses in the Jewellery industry."),
      body("You'll work on meaningful consulting projects — brand strategy, digital systems, training programs, and business transformation — alongside a focused, collaborative team that values clarity, ownership, and excellence."),

      sectionLabel("Open Positions"),
      sectionHeading("Current openings"),
      body("Eyebrow: Open Positions"),
      body("Sub-text: Two roles open right now. If you are passionate about your craft and want to work in a fast-paced consulting environment, we'd love to hear from you."),

      new Paragraph({
        children: [new TextRun({ text: "Position 1: Graphics Designer", bold: true, size: 26, color: NAVY, font: "Calibri" })],
        spacing: { before: 240, after: 80 }
      }),
      body("Tags: Full-time  |  Creative"),
      body("Description: Create compelling visual assets for Jewellery brands — from social media creatives and marketing collaterals to brand identity systems and campaign visuals. Work closely with brand consultants to bring strategy to life through design."),
      body("Experience: 2+ Years"),
      body("Location: On-site / Hybrid"),
      new Paragraph({
        children: [new TextRun({ text: "Skills & Software:", bold: true, size: 21, color: GOLD, font: "Calibri" })],
        spacing: { before: 80, after: 40 }
      }),
      bullet("Adobe Photoshop"), bullet("Adobe Illustrator"), bullet("CorelDRAW"), bullet("Canva"), bullet("Figma"), bullet("Adobe InDesign"), bullet("Video Editing"),
      new Paragraph({
        children: [new TextRun({ text: "Requirements:", bold: true, size: 21, color: GOLD, font: "Calibri" })],
        spacing: { before: 80, after: 40 }
      }),
      bullet("Strong portfolio with brand and marketing design work"),
      bullet("Eye for luxury aesthetics and Jewellery industry visuals"),
      bullet("Able to manage multiple projects and meet deadlines"),
      bullet("Good communication and creative collaboration skills"),
      body("CTA: Apply Now  →  links to contact.html"),

      new Paragraph({
        children: [new TextRun({ text: "Position 2: Software Developer", bold: true, size: 26, color: NAVY, font: "Calibri" })],
        spacing: { before: 280, after: 80 }
      }),
      body("Tags: Full-time  |  Technology"),
      body("Description: Build and maintain web solutions, internal tools, and custom digital platforms for BrillBrains and its Jewellery business clients — from client-facing websites to automation workflows and CRM integrations."),
      body("Experience: 2+ Years"),
      body("Location: On-site / Remote"),
      new Paragraph({
        children: [new TextRun({ text: "Skills & Stack:", bold: true, size: 21, color: GOLD, font: "Calibri" })],
        spacing: { before: 80, after: 40 }
      }),
      bullet("HTML / CSS"), bullet("JavaScript"), bullet("React / Vue"), bullet("Node.js"), bullet("MySQL / SQL"), bullet("Git / GitHub"), bullet("REST APIs"),
      new Paragraph({
        children: [new TextRun({ text: "Requirements:", bold: true, size: 21, color: GOLD, font: "Calibri" })],
        spacing: { before: 80, after: 40 }
      }),
      bullet("Experience building responsive websites and web applications"),
      bullet("Familiar with CRM/ERP integrations and automation tools"),
      bullet("Strong problem-solving and independent working ability"),
      bullet("Comfortable translating business needs into tech solutions"),
      body("CTA: Apply Now  →  links to contact.html"),

      sectionLabel("CTA Band — Open Applications"),
      sectionHeading("Don't see a perfect fit? Reach out anyway"),
      body("Eyebrow: Get in Touch"),
      body("Body: We're always open to meeting talented people. Send us your portfolio or CV and tell us how you'd contribute to the BrillBrains mission."),
      body("Button 1: Send Your CV  →  hello@brillbrainsconsultants.com"),
      body("Button 2: WhatsApp Us  →  +91 73041 92092"),
      divider(),

      // ═══════════════════════════════════════════════
      // PAGE 10 — CONTACT
      // ═══════════════════════════════════════════════
      pageBreak(),
      pageTitle("PAGE 10 — CONTACT US (contact.html)"),

      sectionLabel("Hero Section"),
      sectionHeading("Let's build your empire together"),
      body("Eyebrow: Contact Us"),
      body("Body: Reach out to start a conversation about your Jewellery brand, operations, systems, technology, or team development. We respond within one business day."),
      body("Button 1: WhatsApp Us  →  +91 73041 92092"),
      body("Button 2: Send Email  →  hello@brillbrainsconsultants.com"),

      sectionLabel("Contact Tiles"),
      new Paragraph({
        children: [new TextRun({ text: "Tile 1 — Call", bold: true, size: 22, color: NAVY, font: "Calibri" })],
        spacing: { before: 160, after: 60 }
      }),
      bullet("+91 73041 92092"),
      bullet("+91 73047 92092"),
      new Paragraph({
        children: [new TextRun({ text: "Tile 2 — WhatsApp", bold: true, size: 22, color: NAVY, font: "Calibri" })],
        spacing: { before: 160, after: 60 }
      }),
      bullet("Start WhatsApp Chat"),
      bullet("Best for quick enquiries and consultation requests."),
      new Paragraph({
        children: [new TextRun({ text: "Tile 3 — Email", bold: true, size: 22, color: NAVY, font: "Calibri" })],
        spacing: { before: 160, after: 60 }
      }),
      bullet("hello@brillbrainsconsultants.com"),
      bullet("Share your requirements, business goals, and preferred callback time."),

      sectionLabel("FAQ Section — 8 Questions & Answers"),
      sectionHeading("Frequently asked questions"),
      ...qaItem(
        "What services does BrillBrains offer?",
        "BrillBrains offers four core consulting practices: Brand Consulting (brand strategy, identity, creative services, digital marketing, and video marketing), Business Management Consulting (SOPs, process documentation, system design, and performance tracking), Tech Consulting (CRM/ERP setup, automation, cloud solutions, cybersecurity, and custom software), and professional Training Programs designed specifically for Jewellery industry teams and leaders."
      ),
      ...qaItem(
        "Who is BrillBrains best suited for?",
        "We work with retail jewellers, wholesalers, SME owners, and leadership teams in the Jewellery industry. If you want to move from owner-dependent execution to documented, system-driven, team-powered operations — and build a brand that commands trust and recognition — BrillBrains is the right partner for you."
      ),
      ...qaItem(
        "How does the consulting engagement process work?",
        "Every engagement follows our four-step process: Discover — a deep audit of your brand, operations, identity, and team mindset; Design — tailored strategy, systems, and creative direction for your business; Deploy — rolling out documentation, processes, and training across the organization; and Develop — a long-term partnership focused on continuous improvement, reporting, and scaling."
      ),
      ...qaItem(
        "Do you serve Jewellery businesses outside India?",
        "Yes. BrillBrains serves Jewellery businesses worldwide — retail jewellers and wholesalers across India and internationally. Our consulting, brand, and training programs are designed with the global Jewellery industry in mind, and we work with clients remotely as well as on-site."
      ),
      ...qaItem(
        "What is included in your Brand Consulting service?",
        "Brand Consulting at BrillBrains covers the full spectrum: brand audit, strategy, positioning, voice, identity, and brand guidelines; creative services including logo, social media creatives, print media, and outdoor advertising; content creation including videos, reels, presentations, newsletters, and website copy; digital marketing (SEO, SEM, PPC, Google Ads, WhatsApp marketing, social media management); marketing strategy and campaign planning; video marketing; and The Web Spiders service for domain, hosting, and complete website design and development."
      ),
      ...qaItem(
        "What training programs are available for Jewellery teams?",
        "We offer 10 specialized training programs: Management Training, The Jewel Graduate (Jewellery business foundation skills), Life & Work Culture Transformation, Team Building, Marketing Mastery, Outstanding Orientation (onboarding systems), Digital Shakti (practical technology skills for non-tech staff), Leadership Training, Business Tour Experts (B2B sales travel and buyer meetings), and The Exhibition Excellence (pre/during/post exhibition performance). All programs are customized for Jewellery industry teams."
      ),
      ...qaItem(
        "How is BrillBrains different from a regular marketing agency?",
        "BrillBrains is not a marketing agency — we are a Business Growth Consultant firm. While we do handle brand identity, digital marketing, and creative services, our work goes much deeper: we build operational systems, document processes, set up technology, and develop teams from within. Our goal is to transform the entire business, not just its online presence. We focus on long-term impact, not campaigns."
      ),
      ...qaItem(
        "How do I get started with BrillBrains?",
        "Getting started is simple — reach out via WhatsApp at +91 73041 92092, call us directly at +91 73047 92092, or email hello@brillbrainsconsultants.com. We respond within one business day. Share your requirements, business goals, and a preferred callback time and we will take it from there."
      ),

      sectionLabel("Google Map Section"),
      sectionHeading("Our location"),
      body("Eyebrow: Find Us"),
      body("Location: AGARWAL RESIDENCY, Malad, Adarsh Dugdhalaya, Kailash Nagar, Malad West, Mumbai, Maharashtra 400067"),
      body("(Google Maps embed displayed on website)"),
      divider(),

      // ═══════════════════════════════════════════════
      // FOOTER
      // ═══════════════════════════════════════════════
      pageBreak(),
      pageTitle("FOOTER (All Pages)"),

      sectionLabel("Brand Column"),
      body("Logo: BrillBrains logo"),
      body("Brand Name: BrillBrains"),
      body("Tagline: Strategic & Creative Forever"),
      body("Description: BrillBrains Consultants Pvt. Ltd. strategically transforms Jewellery brands and businesses through consulting, management systems, technology, and professional training."),

      sectionLabel("Explore Column"),
      bullet("About Us  →  about.html"),
      bullet("Our Clients  →  clients.html"),
      bullet("Services  →  services.html"),
      bullet("Careers  →  careers.html"),
      bullet("Contact Us  →  contact.html"),

      sectionLabel("Services Column"),
      bullet("Brand Consulting  →  brand-consulting.html"),
      bullet("Business Management Consulting  →  business-management-consulting.html"),
      bullet("Tech Consulting  →  tech-consulting.html"),
      bullet("Training Programs  →  training-programs.html"),

      sectionLabel("Connect Column"),
      bullet("+91 73041 92092"),
      bullet("+91 73047 92092"),
      bullet("hello@brillbrainsconsultants.com"),

      sectionLabel("Social Media Icons"),
      bullet("Instagram"),
      bullet("Facebook"),
      bullet("X (Twitter)"),
      bullet("LinkedIn"),

      sectionLabel("Footer Bottom Bar"),
      body("© [Year] BrillBrains Consultants Pvt. Ltd. All rights reserved."),

      sectionLabel("Floating WhatsApp Button"),
      body("Chat on WhatsApp  →  +91 73041 92092  (visible on all pages, bottom-right corner)"),
      divider(),

      // End note
      new Paragraph({
        children: [new TextRun({ text: "— End of Document —", italics: true, size: 20, color: GRAY, font: "Calibri" })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 400 }
      })
    ]
  }]
});

const buffer = await Packer.toBuffer(doc);
writeFileSync("BrillBrains-Website-Content.docx", buffer);
console.log("Done: BrillBrains-Website-Content.docx created.");
