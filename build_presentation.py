from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

primary = RGBColor(0x4F, 0x46, 0xE5)
secondary = RGBColor(0xFB, 0xBF, 0x24)
accent = RGBColor(0x10, 0xB9, 0x81)
text_dark = RGBColor(0x17, 0x24, 0x39)

def add_textbox(slide, left, top, width, height, text, font_name="Open Sans", font_size=18, bold=False, color=text_dark):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    run = p.runs[0]
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return tf

slides_data = [
    {
        "module": "Module 1 – Introduction to AI",
        "title": "AI Tools 2025 – For Non-Developers",
        "bullets": [
            "Welcome to a practical, friendly tour of modern AI tools",
            "You’ll learn to ideate, create, and automate without coding",
            "Course roadmap: 10 modules, hands-on demos, final project"
        ],
        "visuals": "Futuristic AI illustration with gradient background and neon circuit overlay",
        "narration": "Welcome to AI Tools 2025 for Non-Developers. Over the next few hours, we’ll explore the tools, workflows, and creative possibilities that help anyone harness AI without writing code.",
        "motion": "Background slow-moving gradient, title slide-in, icon subtle pulse"
    },
    {
        "module": "Module 1 – Introduction to AI",
        "title": "How AI Evolved",
        "bullets": [
            "1950s: Turing Test sparks AI imagination",
            "2010s: Deep learning unlocks image and speech breakthroughs",
            "2020s: Generative AI enables content, code, and creativity"
        ],
        "visuals": "Three-point timeline graphic with milestone icons",
        "narration": "Let’s ground ourselves with a quick history. From Turing’s questions to today’s generative models, AI has steadily become more capable and accessible.",
        "motion": "Timeline elements fade in sequentially"
    },
    {
        "module": "Module 1 – Introduction to AI",
        "title": "Types of AI",
        "bullets": [
            "Narrow AI excels at specific tasks like chatbots and recommenders",
            "General AI is theoretical systems matching human versatility",
            "Super AI would surpass human intelligence—still speculative"
        ],
        "visuals": "Three-column icons for assistant, brain, and rocket",
        "narration": "Most tools we use are narrow AI—specialists. General and super AI remain future concepts, helping us frame what’s possible today versus tomorrow.",
        "motion": "Columns pop up one after another"
    },
    {
        "module": "Module 1 – Introduction to AI",
        "title": "AI Impact on Work and Life",
        "bullets": [
            "Automates repetitive office tasks",
            "Enhances personal productivity with planning and reminders",
            "Unlocks data-driven decisions for small businesses"
        ],
        "visuals": "Split screen of office desk, smartphone, and bar chart",
        "narration": "AI is now an everyday co-worker—drafting emails, summarizing meetings, and surfacing insights so you can focus on higher-value thinking.",
        "motion": "Charts slide from bottom, icons fade in"
    },
    {
        "module": "Module 1 – Introduction to AI",
        "title": "Real-World AI Examples",
        "bullets": [
            "Healthcare: triage chatbots and imaging support",
            "Finance: fraud detection and customer service",
            "Social media: personalized feeds and smart moderation"
        ],
        "visuals": "Industry icons for stethoscope, shield, and hashtag",
        "narration": "Across industries, AI quietly powers operations—from detecting early fraud to curating feeds that keep audiences engaged.",
        "motion": "Icons perform gentle spin-in with text fade"
    },
    {
        "module": "Module 1 – Introduction to AI",
        "title": "Module 1 Summary",
        "bullets": [
            "AI is now practical, not futuristic",
            "Narrow AI powers today’s tools",
            "Impact spans work, creativity, and daily decisions"
        ],
        "visuals": "Checklist graphic with AI glow",
        "narration": "Remember: AI is already delivering value, especially via narrow tools you’ll start using right away.",
        "motion": "Checklist ticks animate sequentially"
    },
    {
        "module": "Module 2 – Essential AI Assistants",
        "title": "Why AI Assistants Matter",
        "bullets": [
            "Speed up research and brainstorming",
            "Offer a 24/7 idea partner",
            "Simplify communication tasks"
        ],
        "visuals": "Chat bubbles orbiting a central assistant icon",
        "narration": "AI assistants extend your brainpower—ideal for drafting, planning, or checking ideas anytime.",
        "motion": "Floating chat bubbles with slow upward motion"
    },
    {
        "module": "Module 2 – Essential AI Assistants",
        "title": "ChatGPT Overview",
        "bullets": [
            "Conversational interface for text, images, and files",
            "Great for drafts, workflows, and explainer scripts",
            "Use system prompts to set tone and role"
        ],
        "visuals": "Chat window mockup with purple accents",
        "narration": "ChatGPT is your generalist partner. Define roles like marketing strategist to control tone and detail.",
        "motion": "Chat bubbles type in from left and right"
    },
    {
        "module": "Module 2 – Essential AI Assistants",
        "title": "Claude by Anthropic",
        "bullets": [
            "Handles long context ideal for large documents",
            "Strong safety guardrails and thoughtful responses",
            "Use for policy drafts and knowledge bases"
        ],
        "visuals": "Document stack with friendly AI avatar",
        "narration": "Claude excels when you need careful reasoning over long documents, making it a favorite for policy and research tasks.",
        "motion": "Document stack slides in, avatar fades"
    },
    {
        "module": "Module 2 – Essential AI Assistants",
        "title": "Google Gemini",
        "bullets": [
            "Multimodal: text, images, and video prompts",
            "Deep integration with Google services",
            "Summarizes Drive files and emails"
        ],
        "visuals": "Icon cluster for camera, doc, and play button",
        "narration": "Gemini shines when you want AI insights across Gmail, Docs, or Slides—everything in one ecosystem.",
        "motion": "Icons rotate with gentle zoom"
    },
    {
        "module": "Module 2 – Essential AI Assistants",
        "title": "Microsoft Copilot",
        "bullets": [
            "Embedded in Word, Excel, and Teams",
            "Drafts content and analyzes spreadsheets",
            "Summarizes meetings automatically"
        ],
        "visuals": "Office suite icons linked to Copilot badge",
        "narration": "Microsoft Copilot lives where your work already happens, so it can draft documents or summarize Teams meetings with just a prompt.",
        "motion": "Connector lines draw in; icons fade"
    },
    {
        "module": "Module 2 – Essential AI Assistants",
        "title": "Module 2 Summary",
        "bullets": [
            "ChatGPT is the versatile creative partner",
            "Claude delivers long-form thoughtful analysis",
            "Gemini and Copilot offer tight productivity integration"
        ],
        "visuals": "Two-by-two comparison grid",
        "narration": "Mix assistants based on context: creativity, depth, or office integration. Use the best co-pilot for each task.",
        "motion": "Grid tiles staggered fade in"
    },
    {
        "module": "Module 3 – AI for Writing and Social",
        "title": "Why Use AI for Content",
        "bullets": [
            "Accelerate ideation and drafting",
            "Maintain brand voice consistently",
            "Personalize messages at scale"
        ],
        "visuals": "Laptop with social icons orbiting",
        "narration": "AI removes blank-page fear and keeps your messaging cohesive across every channel.",
        "motion": "Orbiting icons with slow rotation"
    },
    {
        "module": "Module 3 – AI for Writing and Social",
        "title": "ChatGPT for Copywriting",
        "bullets": [
            "Generate campaign angles, hooks, and calls to action",
            "Use custom instructions for brand tone",
            "Refine drafts with step-by-step prompts"
        ],
        "visuals": "Prompt and result split screen",
        "narration": "Give ChatGPT context—audience, tone, and desired outcome—and iterate. Think of it as a creative copy partner.",
        "motion": "Prompt box slides in left, output right"
    },
    {
        "module": "Module 3 – AI for Writing and Social",
        "title": "Jasper for Marketing Teams",
        "bullets": [
            "Templates for ads, emails, and landing pages",
            "Brand voice library shared across teammates",
            "Campaign collaboration dashboards"
        ],
        "visuals": "Dashboard mockup with brand palette chips",
        "narration": "Jasper scales marketing output by packaging templates and brand voices so teams stay aligned.",
        "motion": "Dashboard zoom-in with highlight sweeps"
    },
    {
        "module": "Module 3 – AI for Writing and Social",
        "title": "Grammarly and Notion AI",
        "bullets": [
            "Grammarly provides tone detector and rewrite suggestions",
            "Notion AI summarizes notes and creates briefs",
            "Combine them for polished, structured documents"
        ],
        "visuals": "Two app cards connected via arrow",
        "narration": "Use Grammarly to refine the language and Notion AI to organize the thinking—quality plus clarity.",
        "motion": "Cards slide from opposite sides and link"
    },
    {
        "module": "Module 3 – AI for Writing and Social",
        "title": "Social Media Applications",
        "bullets": [
            "Draft captions, hashtags, and replies instantly",
            "Repurpose blog posts into threads and carousels",
            "Schedule using AI recommendations"
        ],
        "visuals": "Instagram, TikTok, and Twitter icons with sparkle",
        "narration": "Feed AI your pillar content. Ask for variations tailored to each platform’s style and length.",
        "motion": "Icons pop in sequentially with sparkle"
    },
    {
        "module": "Module 3 – AI for Writing and Social",
        "title": "Blog and Newsletter Workflows",
        "bullets": [
            "Outline, draft, and edit using AI loops",
            "Insert data or quotes manually for authenticity",
            "Use AI to brainstorm titles and metadata"
        ],
        "visuals": "Three-step vertical workflow infographic",
        "narration": "Structure comes first: ask AI for outlines, then expand, then refine. You stay editor-in-chief.",
        "motion": "Steps ascend with fade-up"
    },
    {
        "module": "Module 3 – AI for Writing and Social",
        "title": "Module 3 Summary",
        "bullets": [
            "Templates plus prompts equal efficient writing",
            "Pair ideation tools with editing tools",
            "Always inject brand personality manually"
        ],
        "visuals": "Checklist with pen icon",
        "narration": "Use AI for speed, you for strategy. That balance keeps content authentic.",
        "motion": "Checkmarks animate in order"
    },
    {
        "module": "Module 4 – AI for Images and Logos",
        "title": "Visual Creation Landscape",
        "bullets": [
            "AI art delivers fast mood boards and concepts",
            "Combine AI output with manual polish",
            "Keep brand colors consistent (#4F46E5, #FBBF24, #10B981)"
        ],
        "visuals": "Palette swatches with abstract AI art",
        "narration": "AI visuals jumpstart ideation; you decide final polish to stay on-brand.",
        "motion": "Palette swatches slide in; art fades"
    },
    {
        "module": "Module 4 – AI for Images and Logos",
        "title": "DALL·E for Concept Art",
        "bullets": [
            "Text prompts create unique illustrations",
            "Include camera angles and styles in prompts",
            "Upscale outputs for presentations or posts"
        ],
        "visuals": "Prompt example with resulting artwork",
        "narration": "Prompt with style cues like flat illustration, warm lighting, minimalist icons to guide DALL·E precisely.",
        "motion": "Before and after cards flip"
    },
    {
        "module": "Module 4 – AI for Images and Logos",
        "title": "Canva Magic Media",
        "bullets": [
            "Generate images and edit with drag-and-drop",
            "Brand kits ensure color and font consistency",
            "Quick resizing for multiple formats"
        ],
        "visuals": "Canva interface screenshot with brand kit panel",
        "narration": "Canva bridges AI generation and final layout—perfect for non-designers needing branded assets fast.",
        "motion": "Interface zoom pan"
    },
    {
        "module": "Module 4 – AI for Images and Logos",
        "title": "Adobe Firefly",
        "bullets": [
            "Text-to-image plus generative fill",
            "Licensed training data for commercial safety",
            "Integrates with Photoshop and Express"
        ],
        "visuals": "Firefly logo with Photoshop canvas",
        "narration": "Firefly is ideal when you need safe-to-use visuals and seamless Adobe workflow.",
        "motion": "Logo glow pulse with canvas fade"
    },
    {
        "module": "Module 4 – AI for Images and Logos",
        "title": "Leonardo AI Tips",
        "bullets": [
            "Fine-tune styles and save prompt presets",
            "Great for game, futuristic, stylized looks",
            "Use variations to explore options quickly"
        ],
        "visuals": "Gallery grid with style tags",
        "narration": "Leonardo excels at stylized outputs. Save your favorite settings to generate consistent series.",
        "motion": "Gallery tiles cascade in"
    },
    {
        "module": "Module 4 – AI for Images and Logos",
        "title": "Module 4 Summary",
        "bullets": [
            "Start with AI concepts, finish with brand polish",
            "Choose tools based on workflow needs",
            "Maintain brand kit for consistency"
        ],
        "visuals": "Mood board collage",
        "narration": "Keep prompts, palettes, and edits aligned—AI as your brainstorming partner, not the final designer.",
        "motion": "Collage elements assemble gradually"
    },
    {
        "module": "Module 5 – AI for Video Creation",
        "title": "Why AI for Video",
        "bullets": [
            "Democratizes editing and animation",
            "Reuse scripts across multiple formats",
            "Add AI voiceovers quickly"
        ],
        "visuals": "Play button with waveform and camera icon",
        "narration": "Video no longer requires a studio. AI lets you storyboard, edit, and narrate inside a browser.",
        "motion": "Play button pulses, waveform scrolls"
    },
    {
        "module": "Module 5 – AI for Video Creation",
        "title": "Runway",
        "bullets": [
            "Text-to-video and powerful video editing",
            "Background removal and motion tracking",
            "Ideal for concept teasers"
        ],
        "visuals": "Runway interface screenshot",
        "narration": "Runway’s magic tools let you replace backgrounds or animate concepts with minimal effort.",
        "motion": "Interface elements slide in sequentially"
    },
    {
        "module": "Module 5 – AI for Video Creation",
        "title": "CapCut with AI",
        "bullets": [
            "Templates plus auto captions",
            "Drop in AI voice or TikTok-style edits",
            "Export directly to social platforms"
        ],
        "visuals": "CapCut mobile UI with caption overlay",
        "narration": "CapCut keeps trends within reach—use templates, add your AI-generated script, and publish fast.",
        "motion": "Vertical swipe transition effect"
    },
    {
        "module": "Module 5 – AI for Video Creation",
        "title": "Pika Labs",
        "bullets": [
            "Animate still images with prompts",
            "Cinematic camera motion presets",
            "Great for short loops or ads"
        ],
        "visuals": "Still image transforming into motion",
        "narration": "Pika turns static art into motion. Think cinemagraphs without complex software.",
        "motion": "Image morph demonstration"
    },
    {
        "module": "Module 5 – AI for Video Creation",
        "title": "Canva Video",
        "bullets": [
            "Combine footage, text, and AI-generated clips",
            "Brand kit ensures uniform fonts and colors",
            "Quick resizing for Reels, Shorts, and stories"
        ],
        "visuals": "Canva timeline with brand palette",
        "narration": "Canva Video gives you drag-and-drop editing plus AI media—all within your existing brand kit.",
        "motion": "Timeline scrubber animates across track"
    },
    {
        "module": "Module 5 – AI for Video Creation",
        "title": "Module 5 Summary",
        "bullets": [
            "Mix text-to-video with template editors",
            "Keep scripts, visuals, and voice cohesive",
            "Optimize aspect ratios per platform"
        ],
        "visuals": "Three-device mockup (phone, tablet, laptop)",
        "narration": "Design once, adapt everywhere. AI video tools make omnichannel storytelling manageable.",
        "motion": "Devices slide from center outward"
    },
    {
        "module": "Module 6 – AI for Productivity",
        "title": "Smart Workflows Overview",
        "bullets": [
            "Automate notes, documents, and messages",
            "Reduce meeting fatigue with summaries",
            "Keep teams aligned with AI assistants"
        ],
        "visuals": "Office desk with floating AI icons",
        "narration": "Let AI handle the admin so you can focus on decisions, not documentation.",
        "motion": "Icons hover with subtle bob"
    },
    {
        "module": "Module 6 – AI for Productivity",
        "title": "Notion AI",
        "bullets": [
            "Summarize research and generate action items",
            "Create knowledge bases from notes",
            "Linked databases and AI equal instant updates"
        ],
        "visuals": "Notion workspace mockup",
        "narration": "Use Notion AI to turn messy notes into structured docs and task lists.",
        "motion": "Workspace sections highlight sequentially"
    },
    {
        "module": "Module 6 – AI for Productivity",
        "title": "Microsoft 365 Copilot",
        "bullets": [
            "Drafts emails, slides, and meeting summaries",
            "Generates Excel insights from raw data",
            "Teams recap with action points"
        ],
        "visuals": "Outlook, Word, Excel icons around Copilot badge",
        "narration": "Copilot keeps all your Microsoft workstreams connected and summarized effortlessly.",
        "motion": "Icons orbit around center"
    },
    {
        "module": "Module 6 – AI for Productivity",
        "title": "Google Workspace AI",
        "bullets": [
            "Help me write in Gmail and Docs",
            "Slides auto-layout suggestions",
            "Meet captions plus summaries"
        ],
        "visuals": "Gmail, Docs, and Slides icons with sparkles",
        "narration": "Ask Google AI to draft first versions. Customize from there to keep your voice authentic.",
        "motion": "Icons float in with glow"
    },
    {
        "module": "Module 6 – AI for Productivity",
        "title": "Zoom AI Companion",
        "bullets": [
            "Meeting notes and decisions captured automatically",
            "Smart chat during live sessions",
            "Recaps shared instantly"
        ],
        "visuals": "Zoom meeting mockup with notes panel",
        "narration": "No more frantic note-taking—Zoom AI documents key points in real time.",
        "motion": "Notes panel slides up after main image"
    },
    {
        "module": "Module 6 – AI for Productivity",
        "title": "Module 6 Summary",
        "bullets": [
            "AI assistants keep meetings actionable",
            "Draft, summarize, and analyze faster",
            "Integrate within existing office suites"
        ],
        "visuals": "Calendar with checkmarks",
        "narration": "Adopt AI copilots wherever you already work. It’s the fastest path to ROI.",
        "motion": "Calendar pages flip subtly"
    },
    {
        "module": "Module 7 – AI for Automation",
        "title": "Automation Mindset",
        "bullets": [
            "Map repetitive tasks and automate",
            "Start small with notifications and data syncs",
            "Iterate using metrics and feedback"
        ],
        "visuals": "Flowchart with gears",
        "narration": "Think in workflows: trigger, action, result. AI plus automation keeps teams on autopilot.",
        "motion": "Flow arrows draw in sequence"
    },
    {
        "module": "Module 7 – AI for Automation",
        "title": "Zapier",
        "bullets": [
            "Connect over 6,000 apps",
            "Use AI actions to draft replies and summaries",
            "Build multi-step Zaps for complex flows"
        ],
        "visuals": "App icons connected via Zapier logo",
        "narration": "Zapier is your universal connector—link CRM, email, spreadsheets, and AI for seamless handoffs.",
        "motion": "Connection lines animate"
    },
    {
        "module": "Module 7 – AI for Automation",
        "title": "Make.com",
        "bullets": [
            "Visual scenario builder with drag-and-drop",
            "Branching logic, routers, and delay steps",
            "Great for data-heavy teams"
        ],
        "visuals": "Canvas with connected modules",
        "narration": "Make.com’s canvas lets you visualize every step. Perfect for more advanced automations.",
        "motion": "Modules slide into place with connectors"
    },
    {
        "module": "Module 7 – AI for Automation",
        "title": "IFTTT",
        "bullets": [
            "Simple triggers for IoT, social, and calendars",
            "Great for solo entrepreneurs",
            "Use applets for cross-device sync"
        ],
        "visuals": "Smartphone with IoT icons",
        "narration": "IFTTT is lightweight automation for everyday tasks—ideal when you just need if-this-then-that.",
        "motion": "Icons pop with small bounce"
    },
    {
        "module": "Module 7 – AI for Automation",
        "title": "Trello and Asana AI",
        "bullets": [
            "Summarize boards and task lists",
            "Predict due dates or workload",
            "Automate status updates"
        ],
        "visuals": "Kanban cards with AI sparkle",
        "narration": "Project tools now include AI summaries and smart updates so teams stay aligned without extra meetings.",
        "motion": "Cards slide in; sparkles fade"
    },
    {
        "module": "Module 7 – AI for Automation",
        "title": "Module 7 Summary",
        "bullets": [
            "Automate repetitive steps first",
            "Combine AI text actions with workflows",
            "Monitor results and refine triggers"
        ],
        "visuals": "Automation loop icon",
        "narration": "Automation is iterative—deploy, learn, tweak. Each improvement saves more time.",
        "motion": "Loop icon rotates slowly"
    },
    {
        "module": "Module 8 – No-Code AI Apps",
        "title": "Building Without Code",
        "bullets": [
            "Drag-and-drop builders with AI logic",
            "Prototype internal tools fast",
            "Publish web or mobile apps easily"
        ],
        "visuals": "Tablet displaying app cards",
        "narration": "No-code platforms let you turn ideas into working apps in hours, not weeks.",
        "motion": "Cards fan out with fade"
    },
    {
        "module": "Module 8 – No-Code AI Apps",
        "title": "Glide",
        "bullets": [
            "Transform spreadsheets into mobile apps",
            "AI column for smart responses",
            "Ideal for internal dashboards"
        ],
        "visuals": "Glide interface showing sheet and app preview",
        "narration": "Upload a Google Sheet, map fields, and Glide gives you a polished mobile app with AI-enriched data.",
        "motion": "Spreadsheet transitions into phone mockup"
    },
    {
        "module": "Module 8 – No-Code AI Apps",
        "title": "Softr",
        "bullets": [
            "Build portals and client dashboards",
            "Integrates Airtable and Google Sheets",
            "Add AI blocks for chat or summaries"
        ],
        "visuals": "Web portal mockup with login",
        "narration": "Softr excels at external-facing portals with secure access plus AI-powered insight blocks.",
        "motion": "Sections slide vertically"
    },
    {
        "module": "Module 8 – No-Code AI Apps",
        "title": "Bubble",
        "bullets": [
            "Full web app builder with workflows",
            "Plugin marketplace for AI APIs",
            "Ideal for MVPs and SaaS prototypes"
        ],
        "visuals": "Bubble canvas with workflow panel",
        "narration": "Bubble handles complex logic—use AI plugins to power chatbots or recommendation engines.",
        "motion": "Workflow nodes highlight sequentially"
    },
    {
        "module": "Module 8 – No-Code AI Apps",
        "title": "Vercel and AI",
        "bullets": [
            "Deploy no-code or low-code AI apps",
            "Use Vercel templates for chatbots",
            "Built-in analytics and edge functions"
        ],
        "visuals": "Vercel logo with chatbot widget",
        "narration": "Ship AI apps fast using starter kits, then iterate with built-in analytics on Vercel.",
        "motion": "Widget slides in; metrics tick up"
    },
    {
        "module": "Module 8 – No-Code AI Apps",
        "title": "Module 8 Summary",
        "bullets": [
            "Match platform to use case (internal vs public)",
            "Leverage AI components for personalization",
            "Launch quickly, gather feedback, iterate"
        ],
        "visuals": "Decision tree graphic",
        "narration": "Pick the builder that matches your audience, then let AI features elevate the experience.",
        "motion": "Tree branches grow outward"
    },
    {
        "module": "Module 9 – Bonus Free AI Tools",
        "title": "Value from Free Tools",
        "bullets": [
            "Experiment broadly before upgrading",
            "Combine multiple free tiers",
            "Track usage limits and privacy settings"
        ],
        "visuals": "Gift icons with tool logos",
        "narration": "Free tools are perfect for experimentation—mix and match to cover every workflow.",
        "motion": "Icons drop in like gifts"
    },
    {
        "module": "Module 9 – Bonus Free AI Tools",
        "title": "Tangy and Gamma",
        "bullets": [
            "Tangy acts as AI research assistant with citations",
            "Gamma builds narrative slide decks and docs",
            "Great duo for quick presentations"
        ],
        "visuals": "Split card with two tool interfaces",
        "narration": "Research with Tangy, then turn findings into polished decks with Gamma in minutes.",
        "motion": "Cards flip inward"
    },
    {
        "module": "Module 9 – Bonus Free AI Tools",
        "title": "Tome and Blackbox.ai",
        "bullets": [
            "Tome crafts storytelling decks with AI visuals",
            "Blackbox.ai translates natural language to code",
            "Share snippets with developers even if you don’t code"
        ],
        "visuals": "Tome slide preview next to code editor",
        "narration": "Tome creates dynamic, media-rich decks, while Blackbox generates code snippets you can hand to tech partners.",
        "motion": "Slide pans; code types in"
    },
    {
        "module": "Module 9 – Bonus Free AI Tools",
        "title": "Llama Community Models",
        "bullets": [
            "Open-source chat models like Llama 3",
            "Run locally for privacy",
            "Customize via prompt templates"
        ],
        "visuals": "Laptop with llama icon and terminal",
        "narration": "Explore community-driven models like Llama for private experiments or custom assistants.",
        "motion": "Terminal text scroll effect"
    },
    {
        "module": "Module 9 – Bonus Free AI Tools",
        "title": "Module 9 Summary",
        "bullets": [
            "Free tiers cover research, slides, and code",
            "Watch limits and upgrade when ROI is clear",
            "Mix open models with hosted tools"
        ],
        "visuals": "Budget-friendly badge",
        "narration": "Prototype with free tools first; invest once you prove value.",
        "motion": "Badge shimmer animation"
    },
    {
        "module": "Module 10 – Final Project",
        "title": "Final Project Overview",
        "bullets": [
            "Choose one: social pack, brand kit, chatbot app, or mini site",
            "Apply AI tools end-to-end",
            "Document workflow and lessons learned"
        ],
        "visuals": "Project cards grid",
        "narration": "Your capstone synthesizes everything—pick a project that excites you and showcase your AI-powered process.",
        "motion": "Cards stagger in"
    },
    {
        "module": "Module 10 – Final Project",
        "title": "Project Option: Social Media Pack",
        "bullets": [
            "Plan content pillars and weekly cadence",
            "Use ChatGPT or Jasper for copy",
            "Design visuals or short clips in Canva or Runway"
        ],
        "visuals": "Carousel mockups",
        "narration": "Create a week of social posts with AI copy plus branded visuals, ready to schedule.",
        "motion": "Carousel slides swipe"
    },
    {
        "module": "Module 10 – Final Project",
        "title": "Project Option: Logo and Brand Kit",
        "bullets": [
            "Generate concept directions in DALL·E or Firefly",
            "Transfer final assets into a Canva brand kit",
            "Deliver logo, pattern, and mockups"
        ],
        "visuals": "Logo board with color swatches",
        "narration": "Generate logo directions, then refine and package assets into a cohesive brand kit.",
        "motion": "Swatches fade; logo scales up"
    },
    {
        "module": "Module 10 – Final Project",
        "title": "Project Option: Chatbot App",
        "bullets": [
            "Use Glide or Softr for front-end",
            "Hook ChatGPT API or Llama model",
            "Provide onboarding instructions and FAQs"
        ],
        "visuals": "Chat interface screenshot",
        "narration": "Build a no-code chatbot that answers FAQs or guides users through a process.",
        "motion": "Chat bubbles animate typing"
    },
    {
        "module": "Module 10 – Final Project",
        "title": "Project Option: Mini Website",
        "bullets": [
            "Draft sections with Notion AI or Tome",
            "Build in Bubble or deploy via Vercel",
            "Include contact form and clear CTA"
        ],
        "visuals": "Responsive site mockup",
        "narration": "Launch a simple AI-powered site showcasing services, portfolio, or a lead magnet.",
        "motion": "Desktop-to-mobile responsive animation"
    },
    {
        "module": "Module 10 – Final Project",
        "title": "Execution Checklist",
        "bullets": [
            "Plan: objectives, audience, brand voice",
            "Build: leverage relevant AI tools",
            "Review: gather peer feedback and polish",
            "Showcase: record short walkthrough video"
        ],
        "visuals": "Checklist with icons per step",
        "narration": "Follow this loop—plan, build, review, showcase—to turn your idea into a polished deliverable.",
        "motion": "Icons appear as steps highlight"
    },
    {
        "module": "Module 10 – Final Project",
        "title": "Presenting Your Project",
        "bullets": [
            "Include before-and-after or AI prompt samples",
            "Highlight tool stack and lessons learned",
            "Share metrics or qualitative feedback"
        ],
        "visuals": "Presentation slide mockup",
        "narration": "Tell the story of how AI accelerated your workflow. Show prompts, outputs, and the final impact.",
        "motion": "Slide elements fade from left and right"
    },
    {
        "module": "Module 10 – Final Project",
        "title": "Course Wrap-Up and Next Steps",
        "bullets": [
            "You can ideate, design, and automate with AI",
            "Keep experimenting with new assistants",
            "Join communities and share best practices",
            "Plan ongoing learning or certification"
        ],
        "visuals": "Celebration graphic with confetti in brand colors",
        "narration": "Congrats! You now have a toolkit to build smarter workflows and creative projects with AI. Keep exploring and share your wins.",
        "motion": "Confetti animation with text fade-up"
    }
]

def create_slide(prs, index, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0), prs.slide_width, Inches(0.6)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = primary
    header.line.fill.background()
    header_tf = header.text_frame
    header_tf.text = f"{data['module']}"
    header_p = header_tf.paragraphs[0]
    header_p.font.name = "Montserrat"
    header_p.font.size = Pt(20)
    header_p.font.bold = True
    header_p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    slide_number_box = add_textbox(
        slide,
        Inches(9), Inches(0.05), Inches(1), Inches(0.3),
        f"Slide {index}",
        font_name="Poppins",
        font_size=14,
        bold=True,
        color=RGBColor(0xFF, 0xFF, 0xFF)
    )
    slide_number_box.paragraphs[0].alignment = 2

    title_tf = add_textbox(
        slide,
        Inches(0.5), Inches(0.8), Inches(9), Inches(0.9),
        data["title"],
        font_name="Montserrat",
        font_size=34,
        bold=True,
        color=primary
    )

    bullet_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(5.5), Inches(4.5))
    bullet_tf = bullet_box.text_frame
    bullet_tf.clear()
    for i, bullet in enumerate(data["bullets"]):
        if i == 0:
            p = bullet_tf.paragraphs[0]
        else:
            p = bullet_tf.add_paragraph()
        p.text = bullet
        p.font.name = "Open Sans"
        p.font.size = Pt(22)
        p.font.color.rgb = text_dark
        p.level = 0

    visuals_tf = add_textbox(
        slide,
        Inches(6.2), Inches(1.8), Inches(3.2), Inches(1.6),
        "Suggested Visuals / Icons / Images",
        font_name="Poppins",
        font_size=18,
        bold=True,
        color=secondary
    )
    visuals_para = visuals_tf.add_paragraph()
    visuals_para.text = data["visuals"]
    visuals_para.font.name = "Open Sans"
    visuals_para.font.size = Pt(18)
    visuals_para.font.color.rgb = text_dark

    narration_tf = add_textbox(
        slide,
        Inches(6.2), Inches(3.6), Inches(3.2), Inches(2.0),
        "AI Voice Narration Script",
        font_name="Poppins",
        font_size=18,
        bold=True,
        color=accent
    )
    narration_para = narration_tf.add_paragraph()
    narration_para.text = data["narration"]
    narration_para.font.name = "Open Sans"
    narration_para.font.size = Pt(16)
    narration_para.font.color.rgb = text_dark

    motion_tf = add_textbox(
        slide,
        Inches(6.2), Inches(5.75), Inches(3.2), Inches(1.0),
        "Suggested Motion / Animation",
        font_name="Poppins",
        font_size=18,
        bold=True,
        color=primary
    )
    motion_para = motion_tf.add_paragraph()
    motion_para.text = data["motion"]
    motion_para.font.name = "Open Sans"
    motion_para.font.size = Pt(16)
    motion_para.font.color.rgb = text_dark


def build_presentation(output_path: str):
    prs = Presentation()
    for idx, data in enumerate(slides_data, start=1):
        create_slide(prs, idx, data)
    prs.save(output_path)


if __name__ == "__main__":
    output_file = "AI_Tools_2025_For_Non_Developers.pptx"
    build_presentation(output_file)
    print(f"Presentation saved to {output_file}")
