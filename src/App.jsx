import React, { useEffect, useMemo, useState, useRef } from "react";
import "./index.css";
import pptxgen from "pptxgenjs";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
} from "docx";

const chapters = [
  "Abstract", "Introduction", "Literature Survey", "Problem Statement",
  "Objectives", "Methodology", "System Design", "Implementation",
  "Results", "Conclusion", "References"
];

const starterText = {
  Abstract: "The Smart Classroom Availability System provides a platform for checking classroom availability, managing reservations, and reducing scheduling conflicts.",
  Introduction: "Educational institutions manage classrooms, laboratories, and lecture halls. Manual allocation can cause conflicts. This project proposes a centralized digital solution for classroom availability and reservations.",
  "Literature Survey": "Existing classroom management systems focus on scheduling and resource allocation. Add details from relevant research papers here.",
  "Problem Statement": "Students and faculty can have difficulty finding available classrooms. Manual processes may cause delays, conflicts, and inefficient resource use.",
  Objectives: "Display classroom availability.\nSimplify classroom reservations.\nReduce scheduling conflicts.\nImprove infrastructure utilization.",
  Methodology: "The system uses a web application architecture. Users interact with the frontend, backend services process requests, and a database stores classroom and reservation information.",
  "System Design": "The system contains a user interface, application server, and database. Add your architecture diagram and explain each module here.",
  Implementation: "Describe your actual frontend, backend, database, important modules, and implementation steps.",
  Results: "Describe tested features, screenshots, measurements, and observed results. Include only results you have verified.",
  Conclusion: "The system offers a structured approach to classroom management. Future improvements may include notifications, analytics, mobile support, and automated scheduling.",
  References: "Add the research papers, books, websites, and other academic sources actually used in your project."
};

const seedIdeas = [
  ["AI Code Reviewer", "Analyze source code, identify likely bugs, and suggest improvements.", "AI", ["Python", "Flask", "LLM"]],
  ["Smart Classroom System", "Manage classroom availability, reservations, and timetable conflicts.", "WEB", ["React", "Node.js", "MongoDB"]],
  ["AutoScribe Report Assistant", "Organize project notes and generate structured academic documents.", "AI", ["React", "JavaScript", "DOCX"]],
  ["Campus Lost and Found", "Help students report, search, and claim lost campus items.", "WEB", ["React", "Express", "MongoDB"]],
  ["Student Performance Predictor", "Explore academic data and estimate performance using ML.", "ML", ["Python", "Pandas", "Scikit-learn"]],
  ["Expense Tracker for Students", "Track expenses, set budgets, and visualize spending.", "WEB", ["React", "Chart.js", "LocalStorage"]],
  ["Placement Preparation Portal", "Organize coding practice, aptitude tests, and interview notes.", "WEB", ["React", "Node.js", "MySQL"]],
  ["Phishing URL Detector", "Classify suspicious URLs using explainable security features.", "SEC", ["Python", "ML", "Flask"]],
  ["College Event Manager", "Publish events, manage registrations, and track attendance.", "WEB", ["React", "Firebase", "CSS"]],
  ["AI Resume Feedback Tool", "Check resume structure and suggest improvements based on a target role.", "AI", ["Python", "NLP", "Flask"]],
  ["Library Seat Reservation", "Show available study seats and allow timed reservations.", "WEB", ["React", "Express", "MongoDB"]],
  ["IoT Energy Monitor", "Visualize sensor readings and identify unusual energy usage.", "IOT", ["ESP32", "MQTT", "Python"]],
  ["Accessible Learning Notes", "Organize notes with readable layouts and text-to-speech support.", "WEB", ["React", "Web Speech API", "CSS"]],
  ["Food Waste Donation Portal", "Connect campus food donors with registered collection groups.", "WEB", ["React", "Node.js", "MongoDB"]],
  ["Fake News Literacy Assistant", "Help users inspect claims and record links to credible sources.", "AI", ["Python", "NLP", "Flask"]],
  ["Personalized Study Planner", "Build study schedules from subjects, deadlines, and available time.", "WEB", ["React", "LocalStorage", "Calendar"]],
  ["Medicine Reminder App", "Create reminders and a log of user-entered medication schedules.", "MOBILE", ["React Native", "SQLite"]],
  ["Traffic Sign Classifier", "Train an image classifier for common traffic signs.", "ML", ["Python", "TensorFlow", "OpenCV"]],
  ["Smart Parking Dashboard", "Display parking-slot status and manage reservations.", "IOT", ["ESP32", "React", "Firebase"]],
  ["Online Quiz Platform", "Create quizzes, score attempts, and show topic-wise progress.", "WEB", ["React", "Express", "MySQL"]]
];

const initialPapers = [
  { title: "Smart Classroom Management System", authors: "A. Kumar, S. Rao", year: "2024", source: "IEEE", status: "Reviewed" },
  { title: "IoT Based Classroom Monitoring", authors: "R. Sharma, P. Singh", year: "2023", source: "Springer", status: "Reviewed" },
  { title: "Automated Room Reservation System", authors: "K. Patel", year: "2022", source: "ScienceDirect", status: "Pending" }
];

const readStore = (key, fallback) => {
  try {
    const value = localStorage.getItem(key);
    return value ? JSON.parse(value) : fallback;
  } catch {
    return fallback;
  }
};

function useStoredState(key, initial) {
  const [value, setValue] = useState(() => readStore(key, initial));

  useEffect(() => {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch (error) {
      console.warn("Could not save locally", error);
    }
  }, [key, value]);

  return [value, setValue];
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function downloadText(text, filename, type = "text/plain;charset=utf-8") {
  downloadBlob(new Blob([text], { type }), filename);
}

function esc(value = "") {
  return String(value).replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[c]));
}

function documentText(html) {
  if (typeof document === "undefined") {
    return String(html).replace(/<[^>]+>/g, " ");
  }
  const div = document.createElement("div");
  div.innerHTML = String(html || "");
  return div.innerText || div.textContent || "";
}

/* Convert saved rich-text HTML into Word paragraphs while retaining
   common formatting such as bold, italic, underline, headings and lists. */
function richHtmlToWord(html) {
  const container = document.createElement("div");
  container.innerHTML = String(html || "");
  const output = [];

  function runsFor(node, inherited = {}) {
    if (node.nodeType === Node.TEXT_NODE) {
      if (!node.textContent) return [];
      return [new TextRun({
        text: node.textContent,
        bold: inherited.bold || false,
        italics: inherited.italics || false,
        underline: inherited.underline ? {} : undefined,
        strike: inherited.strike || false
      })];
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return [];

    const tag = node.tagName.toLowerCase();
    const next = {
      bold: inherited.bold || ["b", "strong"].includes(tag),
      italics: inherited.italics || ["i", "em"].includes(tag),
      underline: inherited.underline || tag === "u",
      strike: inherited.strike || tag === "s" || tag === "del"
    };

    return Array.from(node.childNodes).flatMap(child => runsFor(child, next));
  }

  function visit(node) {
    if (node.nodeType !== Node.ELEMENT_NODE) return;

    const tag = node.tagName.toLowerCase();

    if (["h1", "h2", "h3"].includes(tag)) {
      output.push(new Paragraph({
        text: documentText(node.innerHTML),
        heading: tag === "h1"
          ? HeadingLevel.HEADING_1
          : tag === "h2"
            ? HeadingLevel.HEADING_2
            : HeadingLevel.HEADING_3,
        spacing: { before: 240, after: 120 },
        pageBreakBefore: false
      }));
      return;
    }

    if (tag === "ul" || tag === "ol") {
      Array.from(node.children).forEach((li, index) => {
        output.push(new Paragraph({
          children: [
            new TextRun({
              text: tag === "ul" ? "• " : `${index + 1}. `,
              bold: false
            }),
            ...Array.from(li.childNodes).flatMap(child => runsFor(child))
          ],
          indent: { left: 360 },
          spacing: { after: 100 }
        }));
      });
      return;
    }

    if (["p", "div", "blockquote", "pre"].includes(tag)) {
      const runs = Array.from(node.childNodes).flatMap(child => runsFor(child));
      if (runs.length) {
        output.push(new Paragraph({
          children: runs,
          alignment: tag === "blockquote" ? AlignmentType.LEFT : AlignmentType.JUSTIFIED,
          spacing: { after: 160 }
        }));
      }
      return;
    }

    Array.from(node.childNodes).forEach(child => {
      if (child.nodeType === Node.ELEMENT_NODE) visit(child);
      else if (child.nodeType === Node.TEXT_NODE && child.textContent.trim()) {
        output.push(new Paragraph(child.textContent.trim()));
      }
    });
  }

  Array.from(container.childNodes).forEach(node => {
    if (node.nodeType === Node.TEXT_NODE && node.textContent.trim()) {
      output.push(new Paragraph(node.textContent.trim()));
    } else {
      visit(node);
    }
  });

  return output.length ? output : [new Paragraph("No content provided.")];
}


function Login({ onLogin }) {
  const [mode, setMode] = useState("signin");
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [message, setMessage] = useState("");
  const [loading, setLoading] = useState(false);

  async function submit() {
    setMessage("");

    if (mode === "guest") {
      const guest = { name: "Student", email: "" };
      localStorage.setItem(
        "autoscribe-current-user",
        JSON.stringify(guest)
      );
      onLogin(guest);
      return;
    }

    const normalizedEmail = email.trim().toLowerCase();

    if (!normalizedEmail || !password) {
      setMessage("Please enter your email address and password.");
      return;
    }

    if (mode === "create") {
      if (!name.trim()) {
        setMessage("Please enter your full name.");
        return;
      }

      if (password.length < 8 || password.length > 72) {
        setMessage("Password must contain between 8 and 72 characters.");
        return;
      }

      if (password !== confirmPassword) {
        setMessage("Passwords do not match.");
        return;
      }
    }

    setLoading(true);

    try {
      const API_URL = (
        import.meta.env.VITE_API_URL || "http://localhost:5000"
      ).replace(/\/+$/, "");

      const endpoint =
        mode === "create"
          ? "/api/auth/register"
          : "/api/auth/login";

      const payload =
        mode === "create"
          ? {
              name: name.trim(),
              email: normalizedEmail,
              password,
            }
          : {
              email: normalizedEmail,
              password,
            };

      const response = await fetch(`${API_URL}${endpoint}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      const data = await response.json().catch(() => ({}));

      if (!response.ok) {
        throw new Error(
          data.message || "Authentication failed. Please try again."
        );
      }

      if (!data.token || !data.user) {
        throw new Error("The server returned an incomplete response.");
      }

      const current = {
        id: data.user.id,
        name: data.user.name,
        email: data.user.email,
      };

      // Save the token and profile, never the password.
      localStorage.setItem("autoscribe-token", data.token);
      localStorage.setItem(
        "autoscribe-current-user",
        JSON.stringify(current)
      );

      onLogin(current);
    } catch (error) {
      if (error instanceof TypeError) {
        setMessage(
          "Cannot connect to AutoScribe's backend. Make sure the backend is running."
        );
      } else {
        setMessage(error.message || "Something went wrong. Please try again.");
      }
    } finally {
      setLoading(false);
    }
  }

  function switchMode() {
    setMode(mode === "create" ? "signin" : "create");
    setMessage("");
    setPassword("");
    setConfirmPassword("");
  }

  function continueAsGuest() {
    const guest = { name: "Student", email: "" };

    localStorage.setItem(
      "autoscribe-current-user",
      JSON.stringify(guest)
    );

    localStorage.removeItem("autoscribe-token");
    onLogin(guest);
  }

  return (
    <div className="login-page">
      <div className="login-left">
        <div className="brand">
          <div className="brand-mark">✦</div>
          <div>
            <h1>AutoScribe</h1>
            <p>Project Assistant</p>
          </div>
        </div>

        <div className="login-message">
          <span className="eyebrow">B.TECH PROJECT WORKSPACE</span>
          <h2>
            Turn your project
            <br />
            idea into a <span>complete report.</span>
          </h2>
          <p>
            Plan, research, document and prepare your B.Tech project
            from one intelligent workspace.
          </p>

          <div className="login-points">
            <div>✓ Project planning</div>
            <div>✓ Literature survey</div>
            <div>✓ Documentation</div>
            <div>✓ Report &amp; viva preparation</div>
          </div>
        </div>
      </div>

      <div className="login-right">
        <div className="login-card">
          <div className="login-card-icon">✦</div>

          <h2>
            {mode === "create" ? "Create your account" : "Welcome back"}
          </h2>

          <p>
            {mode === "create"
              ? "Create an account to organize your project."
              : "Sign in to continue to AutoScribe."}
          </p>

          {mode === "create" && (
            <>
              <label>Full name</label>
              <input
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="Your name"
                autoComplete="name"
                maxLength={80}
                disabled={loading}
              />
            </>
          )}

          <label>Email address</label>
          <input
            type="email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            placeholder="student@example.com"
            autoComplete="email"
            disabled={loading}
          />

          <label>Password</label>
          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            placeholder="Enter your password"
            autoComplete={
              mode === "create" ? "new-password" : "current-password"
            }
            disabled={loading}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !loading) submit();
            }}
          />

          {mode === "create" && (
            <>
              <label>Confirm password</label>
              <input
                type="password"
                value={confirmPassword}
                onChange={(e) => setConfirmPassword(e.target.value)}
                placeholder="Re-enter your password"
                autoComplete="new-password"
                disabled={loading}
                onKeyDown={(e) => {
                  if (e.key === "Enter" && !loading) submit();
                }}
              />
            </>
          )}

          <button
            className="primary-button"
            onClick={submit}
            disabled={loading}
          >
            {loading
              ? "Please wait..."
              : mode === "create"
                ? "Create account →"
                : "Sign in →"}
          </button>

          <div className="login-divider">
            <span>or</span>
          </div>

          <button
            className="secondary-button"
            onClick={switchMode}
            disabled={loading}
          >
            {mode === "create" ? "Back to sign in" : "Create New Account"}
          </button>

          <button
            className="secondary-button"
            style={{ marginTop: 8 }}
            onClick={continueAsGuest}
            disabled={loading}
          >
            Continue as Guest
          </button>

          {message && (
            <p role="alert" className="auth-error">
              {message}
            </p>
          )}

          <small className="login-note">
            Your account is stored in the AutoScribe database.
            Passwords are securely hashed by the server.
          </small>
        </div>
      </div>
    </div>
  );
}

function PageHeader({ icon, title, description, children }) {
  return <>
    <div className="page-heading">
      <div className="page-icon">{icon}</div>
      <div><span className="eyebrow">AUTOSCRIBE MODULE</span><h1>{title}</h1><p>{description}</p></div>
    </div>
    {children}
  </>;
}

function Field({ label, children }) {
  return <div className="field"><label>{label}</label>{children}</div>;
}

function TopBar({ user }) {
  return <header className="topbar">
    <div><span className="eyebrow">AUTOSCRIBE WORKSPACE</span><h1>Good to see you, {user?.name || "Student"} 👋</h1><p>Let's continue building your project.</p></div>
    <div className="profile"><div className="avatar">{(user?.name || "S").charAt(0).toUpperCase()}</div><div><strong>{user?.name || "Student"}</strong><small>B.Tech CSE</small></div></div>
  </header>;
}

function ActionCard({ icon, title, text, onClick }) {
  return <div className="action-card"><div className="action-icon">{icon}</div><h3>{title}</h3><p>{text}</p><button onClick={onClick}>Open →</button></div>;
}

function DashboardHome({ setActive, user }) {
  return <>
    <TopBar user={user} />
    <section className="hero-project">
      <div><span className="status-badge">IN PROGRESS</span><h2>Smart Classroom Availability System</h2><p>Continue building your project documentation and final report.</p></div>
      <button className="primary-button small" onClick={() => setActive("Documentation")}>Continue Project →</button>
    </section>
    <section className="progress-card">
      <div className="section-heading"><div><h3>Project Progress</h3><p>Keep working through each stage</p></div></div>
      <div className="progress-track"><div style={{ width: "62%" }} /></div>
      <div className="progress-steps"><span className="done">✓ Idea</span><span className="done">✓ Proposal</span><span className="done">✓ Literature</span><span>Documentation</span><span>Report</span><span>PPT &amp; Viva</span></div>
    </section>
    <div className="section-title-row"><h2>Quick Actions</h2></div>
    <section className="action-grid">
      <ActionCard icon="💡" title="Project Ideas" text="Generate and save project concepts." onClick={() => setActive("Project Ideas")} />
      <ActionCard icon="📋" title="Proposal" text="Create and save your project proposal." onClick={() => setActive("Proposal & Planning")} />
      <ActionCard icon="📚" title="Literature Survey" text="Manage research papers and references." onClick={() => setActive("Literature Survey")} />
      <ActionCard icon="📝" title="Documentation" text="Format and save chapters." onClick={() => setActive("Documentation")} />
      <ActionCard icon="📄" title="Report Builder" text="Build and download your report." onClick={() => setActive("Report Builder")} />
      <ActionCard icon="🎤" title="PPT & Viva" text="Create a PowerPoint and practice viva." onClick={() => setActive("PPT & Viva")} />
    </section>
  </>;
}

function ProjectIdeas({ onUseIdea }) {
  const [branch, setBranch] = useState("Computer Science & Engineering");
  const [type, setType] = useState("Any");
  const [difficulty, setDifficulty] = useState("Intermediate");
  const [technology, setTechnology] = useState("");
  const [count, setCount] = useState(8);
  const [batch, setBatch] = useState(0);
  const [generated, setGenerated] = useState(false);

  const ideas = useMemo(() => {
    let list = [...seedIdeas];
    if (type === "Web Application") list = list.filter(x => x[3].some(t => ["React", "Node.js", "Express", "MongoDB", "Firebase", "CSS", "MySQL"].includes(t)));
    if (type === "AI / Machine Learning") list = list.filter(x => ["AI", "ML", "SEC"].includes(x[2]));
    if (type === "Mobile Application") list = list.filter(x => x[3].some(t => t.includes("Native")));
    if (type === "IoT") list = list.filter(x => x[2] === "IOT");
    if (difficulty === "Beginner") list = list.filter(x => !["Traffic Sign Classifier", "IoT Energy Monitor", "Phishing URL Detector"].includes(x[0]));
    if (difficulty === "Advanced") list = list.filter(x => ["Traffic Sign Classifier", "IoT Energy Monitor", "Phishing URL Detector", "Student Performance Predictor", "Fake News Literacy Assistant", "Smart Parking Dashboard"].includes(x[0]));
    const terms = technology.trim().toLowerCase().split(/[,; ]+/).filter(Boolean);
    if (terms.length) {
      const matched = list.filter(x => x[3].some(t => terms.some(term => t.toLowerCase().includes(term))));
      if (matched.length) list = matched;
    }
    return list.length ? list : seedIdeas;
  }, [type, difficulty, technology]);

  const shown = useMemo(() => {
    const offset = (batch * count) % ideas.length;
    return Array.from({ length: Math.min(count, ideas.length) }, (_, i) => ideas[(offset + i) % ideas.length]);
  }, [ideas, batch, count]);

  return <PageHeader icon="💡" title="Project Ideas" description="Discover project concepts that match your interests.">
    <div className="generator-card">
      <div className="generator-heading"><div><span className="eyebrow">PROJECT IDEA GENERATOR</span><h2>What do you want to build?</h2><p>Choose preferences and generate ideas.</p></div><div className="ai-circle">✦</div></div>
      <div className="form-grid">
        <Field label="Branch"><select value={branch} onChange={e => setBranch(e.target.value)}><option>Computer Science &amp; Engineering</option><option>Artificial Intelligence</option><option>Information Technology</option><option>Electronics &amp; Communication</option></select></Field>
        <Field label="Project Type"><select value={type} onChange={e => setType(e.target.value)}><option>Any</option><option>Web Application</option><option>AI / Machine Learning</option><option>Mobile Application</option><option>IoT</option></select></Field>
        <Field label="Difficulty"><select value={difficulty} onChange={e => setDifficulty(e.target.value)}><option>Beginner</option><option>Intermediate</option><option>Advanced</option></select></Field>
        <Field label="Technology"><input value={technology} onChange={e => setTechnology(e.target.value)} placeholder="React, Python, MongoDB" /></Field>
        <Field label="Ideas per generation"><select value={count} onChange={e => setCount(Number(e.target.value))}><option value={6}>6 ideas</option><option value={8}>8 ideas</option><option value={12}>12 ideas</option><option value={16}>16 ideas</option><option value={20}>20 ideas</option></select></Field>
      </div>
      <button className="primary-button" onClick={() => { setBatch(0); setGenerated(true); }}>✦ Generate Ideas</button>
    </div>
    <div className="section-title-row"><h2>{generated ? "Generated Project Ideas" : "Suggested Projects"}</h2><button className="secondary-button" onClick={() => { setBatch(b => b + 1); setGenerated(true); }}>Generate More →</button></div>
    <p>{branch} · {difficulty} · Showing {shown.length} ideas.</p>
    <div className="idea-grid">{shown.map((idea, i) => <div className="idea-card" key={`${idea[0]}-${batch}-${i}`}><span className="idea-tag">{idea[2]}</span><h3>{idea[0]}</h3><p>{idea[1]}</p><div className="tag-list">{idea[3].map(t => <span key={t}>{t}</span>)}</div><button onClick={() => onUseIdea(idea)}>Use This Idea →</button></div>)}</div>
  </PageHeader>;
}

function ProposalPlanning({ proposal, setProposal, savedPrompts, setSavedPrompts }) {
  const [draft, setDraft] = useState(proposal);
  const [activeSaved, setActiveSaved] = useState(null);
  useEffect(() => setDraft(proposal), [proposal]);

  function save() {
    const item = { ...draft, id: activeSaved || `${Date.now()}`, savedAt: new Date().toLocaleString() };
    setProposal(draft);
    setSavedPrompts(old => [item, ...old.filter(x => x.id !== item.id)]);
    setActiveSaved(item.id);
  }

  function load(item) {
    setDraft(item);
    setProposal(item);
    setActiveSaved(item.id);
  }

  return <PageHeader icon="📋" title="Proposal & Planning" description="Write project details and save multiple proposal drafts.">
    <div className="two-column">
      <div className="white-card form-card">
        <div className="card-heading"><div><h2>Project Details</h2><p>Save as many drafts as you need.</p></div></div>
        <Field label="Project Title"><input value={draft.title || ""} onChange={e => setDraft({ ...draft, title: e.target.value })} /></Field>
        <Field label="Problem Statement"><textarea value={draft.problem || ""} onChange={e => setDraft({ ...draft, problem: e.target.value })} /></Field>
        <Field label="Objectives"><textarea value={draft.objectives || ""} onChange={e => setDraft({ ...draft, objectives: e.target.value })} /></Field>
        <Field label="Proposed Solution"><textarea value={draft.solution || ""} onChange={e => setDraft({ ...draft, solution: e.target.value })} /></Field>
        <Field label="Technologies"><input value={draft.technologies || ""} onChange={e => setDraft({ ...draft, technologies: e.target.value })} /></Field>
        <button className="primary-button" onClick={save}>Save Proposal / Prompt</button>
        <button className="secondary-button" style={{ marginTop: 8 }} onClick={() => { setDraft({ title: "", problem: "", objectives: "", solution: "", technologies: "" }); setActiveSaved(null); }}>New Proposal</button>
        <button className="secondary-button" style={{ marginTop: 8 }} onClick={() => downloadText(JSON.stringify(draft, null, 2), "autoscribe-proposal.json", "application/json")}>Export Proposal</button>
      </div>
      <div className="white-card preview-card">
        <div className="preview-title"><h2>Saved Prompts</h2><span>{savedPrompts.length} SAVED</span></div>
        {savedPrompts.length === 0 ? <p>Your saved prompts appear here.</p> : savedPrompts.map(item => <div className="saved-prompt" key={item.id} style={{ padding: "12px 0", borderBottom: "1px solid #e6ebf2" }}>
          <strong>{item.title || "Untitled proposal"}</strong><small style={{ display: "block" }}>{item.savedAt || "Saved locally"}</small>
          <div className="button-row" style={{ marginTop: 8 }}>
            <button className="secondary-button" onClick={() => load(item)}>Open</button>
            <button className="secondary-button" onClick={() => { const title = window.prompt("Rename proposal", item.title || "Untitled proposal"); if (title !== null) setSavedPrompts(old => old.map(x => x.id === item.id ? { ...x, title } : x)); }}>Rename</button>
            <button className="secondary-button" onClick={() => { if (window.confirm("Delete this saved proposal?")) setSavedPrompts(old => old.filter(x => x.id !== item.id)); }}>Delete</button>
          </div>
        </div>)}
      </div>
    </div>
  </PageHeader>;
}

function LiteratureSurvey({ papers, setPapers }) {
  const [show, setShow] = useState(false);
  const [search, setSearch] = useState("");
  const [form, setForm] = useState({ title: "", authors: "", year: "", source: "" });

  function add() {
    if (!Object.values(form).every(x => x.trim())) {
      alert("Please fill all fields.");
      return;
    }
    setPapers(old => [...old, { ...form, status: "Pending", id: Date.now() }]);
    setForm({ title: "", authors: "", year: "", source: "" });
    setShow(false);
  }

  const filtered = papers.filter(p => `${p.title} ${p.authors} ${p.source}`.toLowerCase().includes(search.toLowerCase()));

  return <PageHeader icon="📚" title="Literature Survey" description="Collect and organize the research you actually reviewed.">
    <div className="stats-grid">
      <div className="stat-card"><strong>{papers.length}</strong><span>Research Papers</span></div>
      <div className="stat-card"><strong>{papers.filter(p => p.status === "Reviewed").length}</strong><span>Reviewed</span></div>
      <div className="stat-card"><strong>{papers.filter(p => p.status !== "Reviewed").length}</strong><span>Pending</span></div>
      <button className="primary-button" onClick={() => setShow(!show)}>+ Add Research Paper</button>
    </div>
    {show && <div className="white-card form-card"><h2>Add Research Paper</h2><div className="form-grid">{[["title", "Paper Title"], ["authors", "Authors"], ["year", "Year"], ["source", "Journal / Conference"]].map(([key, label]) => <Field label={label} key={key}><input value={form[key]} onChange={e => setForm({ ...form, [key]: e.target.value })} placeholder={label} /></Field>)}</div><div className="button-row"><button className="secondary-button" onClick={() => setShow(false)}>Cancel</button><button className="primary-button" onClick={add}>Add Paper</button></div></div>}
    <div className="white-card">
      <div className="card-heading"><h2>Research Papers</h2><input className="search" value={search} onChange={e => setSearch(e.target.value)} placeholder="Search papers..." /></div>
      <div className="table"><div className="table-head"><span>PAPER</span><span>AUTHORS</span><span>YEAR</span><span>SOURCE</span><span>STATUS</span></div>
        {filtered.map((p, i) => <div className="table-row" key={p.id || i}><div><strong>{p.title}</strong><small>Research Paper</small></div><span>{p.authors}</span><span>{p.year}</span><span>{p.source}</span><button className="secondary-button" onClick={() => setPapers(old => old.map(x => x === p ? { ...x, status: x.status === "Reviewed" ? "Pending" : "Reviewed" } : x))}>{p.status} · toggle</button><button className="secondary-button" onClick={() => setPapers(old => old.filter(x => x !== p))}>Delete</button></div>)}
      </div>
    </div>
  </PageHeader>;
}


function Documentation({ docs, setDocs }) {
 const [chapter, setChapter] = useState("Abstract");
const [content, setContent] = useState("");
const editorRef = useRef(null);

  const chapters = [
    "Abstract",
    "Introduction",
    "Literature Survey",
    "Problem Statement",
    "Objectives",
    "Methodology",
    "System Design",
    "Implementation",
    "Results",
    "Conclusion",
    "References",
  ];

useEffect(() => {
    const savedContent = docs?.[chapter] || "";
    setContent(savedContent);

    if (editorRef.current) {
      editorRef.current.innerHTML = savedContent;
    }
  }, [chapter, docs]);

  const saveContent = (html) => {
    setContent(html);
    setDocs((previous) => ({
      ...previous,
      [chapter]: html,
    }));
  };

  const runCommand = (command, value = null) => {
    const editor = editorRef.current;
    if (!editor) return;

    editor.focus();

    try {
      document.execCommand(command, false, value);
      saveContent(editor.innerHTML);
    } catch (error) {
      console.error("Formatting error:", error);
    }
  };

  const insertHeading = (tag) => {
    runCommand("formatBlock", tag);
  };

  const handleInput = (event) => {
    saveContent(event.currentTarget.innerHTML);
  };

  return (
    <div className="documentation-page">
      <div className="documentation-header">
        <div>
          <h2>Documentation</h2>
          <p>Write, format, and save your project documentation.</p>
        </div>

        <button
          type="button"
          onClick={() => {
            const editor = editorRef.current;
            if (editor) saveContent(editor.innerHTML);
          }}
        >
          Save Chapter
        </button>
      </div>

      <div className="documentation-chapters">
        {chapters.map((item) => (
          <button
            type="button"
            key={item}
            className={chapter === item ? "active" : ""}
            onClick={() => {
              if (editorRef.current) {
                saveContent(editorRef.current.innerHTML);
              }
              setChapter(item);
            }}
          >
            {item}
          </button>
        ))}
      </div>

      <div className="editor-tools">
        <button type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => runCommand("bold")}>
          <b>B</b>
        </button>
        <button type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => runCommand("italic")}>
          <i>I</i>
        </button>
        <button type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => runCommand("underline")}>
          <u>U</u>
        </button>

        <button type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => insertHeading("h1")}>
          Heading 1
        </button>
        <button type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => insertHeading("h2")}>
          Heading 2
        </button>
        <button type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => runCommand("insertUnorderedList")}>
          • Bullets
        </button>
        <button type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => runCommand("insertOrderedList")}>
          1. Numbering
        </button>
        <button type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => runCommand("justifyLeft")}>
          Align Left
        </button>
        <button type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => runCommand("justifyCenter")}>
          Center
        </button>
        <button type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => runCommand("removeFormat")}>
          Clear Formatting
        </button>
      </div>

      <div
        ref={editorRef}
        className="rich-editor"
        contentEditable
        suppressContentEditableWarning
        onInput={handleInput}
        onBlur={(event) => saveContent(event.currentTarget.innerHTML)}
        role="textbox"
        aria-label={`${chapter} editor`}
        aria-multiline="true"
        data-placeholder={`Start writing your ${chapter.toLowerCase()} here...`}
      />

      <p className="documentation-save-status">
        Editing: <strong>{chapter}</strong> · Changes are saved as you type.
      </p>
    </div>
  );
}

function ReportBuilder({ proposal, docs }) {
  const [title, setTitle] = useState(proposal.title || "Smart Classroom Availability System");
  const [student, setStudent] = useState("");
  const [guide, setGuide] = useState("");
  const [department, setDepartment] = useState("Computer Science & Engineering");
  const [abstract, setAbstract] = useState(documentText(docs.Abstract || starterText.Abstract));
  const [message, setMessage] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => setTitle(proposal.title || "Smart Classroom Availability System"), [proposal.title]);
  useEffect(() => setAbstract(documentText(docs.Abstract || starterText.Abstract)), [docs.Abstract]);

  const sections = [
    ["Abstract", abstract],
    ...chapters.filter(c => c !== "Abstract").map(c => [c, docs[c] || starterText[c] || "Add your project-specific content."])
  ];

  async function generate() {
    if (busy) return;
    if (!title.trim()) return setMessage("Please enter a project title.");

    setBusy(true);
    setMessage("");

    try {
      const children = [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 2200, after: 400 }, children: [new TextRun({ text: "PROJECT REPORT", bold: true, size: 36, color: "123D70" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400, after: 400 }, children: [new TextRun({ text: title.trim(), bold: true, size: 32, color: "123D70" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "A B.Tech Project Report", size: 24 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 800 }, children: [new TextRun({ text: `Submitted by: ${student || "Student Name"}`, size: 24 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `Project Guide: ${guide || "Guide Name"}`, size: 24 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: department, size: 24 })] }),
        new Paragraph({ children: [], pageBreakBefore: true }),
        new Paragraph({ text: "Declaration", heading: HeadingLevel.HEADING_1 }),
        new Paragraph("I declare that this report describes the work carried out for the project. Adapt this statement to your institution's required format."),
        new Paragraph({ text: "Acknowledgement", heading: HeadingLevel.HEADING_1, spacing: { before: 300 } }),
        new Paragraph("I express my gratitude to my project guide, department, and everyone who supported this work. Edit this acknowledgement to reflect actual contributions.")
      ];

      sections.forEach(([heading, content]) => {
        children.push(new Paragraph({
          text: heading,
          heading: HeadingLevel.HEADING_1,
          pageBreakBefore: true,
          spacing: { before: 240, after: 200 }
        }));

        const html = String(content || "");
        const paragraphs = richHtmlToWord(html);
        children.push(...paragraphs);
      });

      const report = new Document({
        creator: "AutoScribe",
        title: title.trim(),
        description: "B.Tech project report",
        sections: [{ children }]
      });

      const blob = await Packer.toBlob(report);
      const filename = (title.trim().replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase() || "project-report") + "-final-report.docx";
      downloadBlob(blob, filename);
      setMessage("Your Word report was generated and downloaded. Open it in Word or WPS Writer to review it.");
    } catch (error) {
      console.error(error);
      setMessage("Report generation failed. Check the terminal for the error.");
    } finally {
      setBusy(false);
    }
  }

  return <PageHeader icon="📄" title="Report Builder" description="Create a Word report from your project details and saved chapters.">
    <div className="report-layout">
      <div className="white-card form-card">
        <span className="eyebrow">REPORT SETTINGS</span><h2>Build Final Report</h2>
        <Field label="Project Title"><input value={title} onChange={e => setTitle(e.target.value)} /></Field>
        <Field label="Student Name"><input value={student} onChange={e => setStudent(e.target.value)} placeholder="Enter student name" /></Field>
        <Field label="Guide Name"><input value={guide} onChange={e => setGuide(e.target.value)} placeholder="Enter guide name" /></Field>
        <Field label="Department"><input value={department} onChange={e => setDepartment(e.target.value)} /></Field>
        <Field label="Abstract"><textarea value={abstract} onChange={e => setAbstract(e.target.value)} /></Field>
        <button className="primary-button full" onClick={generate} disabled={busy}>{busy ? "Generating..." : "⬇ Download Final Report (.docx)"}</button>
        {message && <div className="success-message" role="status">{message}</div>}
      </div>
      <div className="report-preview"><div className="report-cover"><span>PROJECT REPORT</span><h1>{title || "Your Project Title"}</h1><p>A B.Tech Project Report</p><div className="cover-line" /><small>{student || "Student Name"} · {department}</small><small>Guide: {guide || "Guide Name"}</small></div></div>
    </div>
    <div className="report-sections">{["Cover Page", "Declaration", "Acknowledgement", ...chapters].map((item, i) => <div className="report-section" key={item}><span>{String(i + 1).padStart(2, "0")}</span><strong>{item}</strong><small>{item === "Abstract" ? "Included" : docs[item] ? "Saved chapter" : "Sample / editable"}</small></div>)}</div>
    <p>Replace all sample content with your actual project work before submission.</p>
  </PageHeader>;
}

const vivaBank = [
  ["What is the main objective of your project?", "Explain the specific problem, intended users, and intended outcome."],
  ["Why did you choose this technology stack?", "Explain how each technology fits your requirements and implementation."],
  ["What problem does your system solve?", "Describe the current process, its limitations, and your proposed improvement."],
  ["Explain your project architecture.", "Describe the frontend, backend, database, and how requests flow through them."],
  ["What are the main modules?", "List the actual modules and explain each module's responsibility."],
  ["How did you test the application?", "Describe the test cases you actually ran and their observed results."],
  ["What challenges did you face?", "Explain a real development problem and how you solved it."],
  ["What are the limitations?", "Describe the current limitations and features that still need improvement."],
  ["What is the role of the database?", "Explain the data stored and how the application accesses it."],
  ["What are your future enhancements?", "Describe realistic improvements such as accessibility, deployment, or monitoring."]
];

function PPTViva({ proposal, docs }) {
  const [title, setTitle] = useState(proposal.title || "Smart Classroom Availability System");
  const [slides, setSlides] = useState(10);
  const [level, setLevel] = useState("All levels");
  const [questions, setQuestions] = useState(vivaBank);
  const [message, setMessage] = useState("");

  useEffect(() => setTitle(proposal.title || "Smart Classroom Availability System"), [proposal.title]);

  const slideContent = [
    ["Title", title, "B.Tech Project Presentation"],
    ["Problem Statement", documentText(docs["Problem Statement"] || starterText["Problem Statement"])],
    ["Objectives", documentText(docs.Objectives || starterText.Objectives)],
    ["Literature Survey", documentText(docs["Literature Survey"] || starterText["Literature Survey"])],
    ["Proposed System", documentText(docs.Introduction || starterText.Introduction)],
    ["System Design", documentText(docs["System Design"] || starterText["System Design"])],
    ["Methodology", documentText(docs.Methodology || starterText.Methodology)],
    ["Implementation", documentText(docs.Implementation || starterText.Implementation)],
    ["Results", documentText(docs.Results || starterText.Results)],
    ["Conclusion", documentText(docs.Conclusion || starterText.Conclusion)],
    ["References", documentText(docs.References || starterText.References)]
  ];

  async function makePpt() {
    try {
      const pptx = new pptxgen();
      pptx.layout = "LAYOUT_WIDE";
      pptx.author = "AutoScribe";
      pptx.subject = "B.Tech project presentation";
      pptx.title = title;
      pptx.theme = { headFontFace: "Aptos Display", bodyFontFace: "Aptos", lang: "en-US" };

      const picked = slideContent.slice(0, Math.min(Number(slides), slideContent.length));

      picked.forEach(([heading, body], i) => {
        const slide = pptx.addSlide();
        slide.background = { color: i === 0 ? "123D70" : "F5F8FC" };

        slide.addText(heading, {
          x: 0.65, y: 0.45, w: 12, h: 0.65,
          fontFace: "Aptos Display", fontSize: i === 0 ? 30 : 26,
          bold: true, color: i === 0 ? "FFFFFF" : "123D70",
          margin: 0.05
        });

        slide.addShape(pptx.ShapeType.line, {
          x: 0.7, y: 1.25, w: 11.9, h: 0,
          line: { color: "2C80D3", width: 2 }
        });

        slide.addText(String(body || "Add project-specific details here."), {
          x: 0.8, y: 1.65, w: 11.6, h: 4.8,
          fontFace: "Aptos", fontSize: 18,
          color: i === 0 ? "FFFFFF" : "25344A",
          breakLine: false, fit: "shrink", valign: "top", margin: 0.12
        });

        slide.addText(String(i + 1), {
          x: 12.2, y: 7.05, w: 0.4, h: 0.2,
          fontSize: 10, color: i === 0 ? "D9EAFE" : "65758B",
          align: "right"
        });
      });

      await pptx.writeFile({
        fileName: `${title.replace(/[^a-z0-9]+/gi, "-").toLowerCase() || "project-presentation"}.pptx`
      });
      setMessage("PowerPoint created and downloaded.");
    } catch (error) {
      console.error(error);
      setMessage("PowerPoint generation failed. Check the terminal.");
    }
  }

  function makeQuestions() {
    setQuestions(
      level === "Basic" ? vivaBank.slice(0, 4) :
      level === "Advanced" ? vivaBank.slice(4) :
      vivaBank
    );
  }

  return <PageHeader icon="🎤" title="PPT & Viva" description="Generate an editable PowerPoint and practice viva questions.">
    <div className="ppt-grid">
      <div className="white-card form-card">
        <span className="eyebrow">PRESENTATION</span><h2>Generate Actual PPT</h2>
        <Field label="Presentation Title"><input value={title} onChange={e => setTitle(e.target.value)} /></Field>
        <Field label="Number of Slides"><select value={slides} onChange={e => setSlides(Number(e.target.value))}><option value={8}>8 Slides</option><option value={9}>9 Slides</option><option value={10}>10 Slides</option><option value={11}>11 Slides</option></select></Field>
        <button className="primary-button full" onClick={makePpt}>Download {slides}-Slide PPTX →</button>
      </div>
      <div className="white-card form-card">
        <span className="eyebrow">VIVA PREPARATION</span><h2>Viva Questions &amp; Answers</h2>
        <Field label="Question level"><select value={level} onChange={e => setLevel(e.target.value)}><option>All levels</option><option>Basic</option><option>Advanced</option></select></Field>
        <button className="primary-button" onClick={makeQuestions}>Generate Viva Questions</button>
        <button className="secondary-button" style={{ marginTop: 8 }} onClick={() => downloadText(questions.map((q, i) => `${i + 1}. ${q[0]}\nAnswer: ${q[1]}`).join("\n\n"), "autoscribe-viva-questions.txt")}>Download Q&amp;A</button>
      </div>
    </div>
    {message && <div className="success-message">{message}</div>}
    <div className="white-card" style={{ marginTop: 20 }}><h2>Practice Viva Questions</h2>{questions.map((q, i) => <details className="question-row" key={q[0]} style={{ display: "block", padding: "12px 0" }}><summary><strong>{i + 1}. {q[0]}</strong></summary><p>{q[1]}</p></details>)}</div>
  </PageHeader>;
}

function Validation({ proposal, docs }) {
  const checks = [
    ["Project title completed", !!proposal.title],
    ["Problem statement completed", !!proposal.problem],
    ["Objectives completed", !!proposal.objectives],
    ["Proposal saved", !!proposal.savedAt],
    ["Abstract chapter exists", !!docs.Abstract],
    ["References added", !!docs.References]
  ];
  const done = checks.filter(x => x[1]).length;

  return <PageHeader icon="✓" title="Validation" description="Check project readiness based on saved work.">
    <div className="validation-summary"><div><span className="eyebrow">PROJECT READINESS</span><h2>{Math.round(done / checks.length * 100)}% Complete</h2><p>{done} of {checks.length} checks completed.</p></div></div>
    <div className="white-card validation-card">{checks.map(([label, ok]) => <div className="validation-row" key={label}><div className={ok ? "check-circle checked" : "check-circle"}>{ok ? "✓" : ""}</div><div><strong>{label}</strong><small>{ok ? "Completed" : "Needs attention"}</small></div><span className={ok ? "valid-text" : "warning-text"}>{ok ? "Ready" : "Pending"}</span></div>)}</div>
  </PageHeader>;
}

function ProgressTracker({ proposal, docs, papers }) {
  const stages = [
    ["Project Ideas", true],
    ["Proposal & Planning", !!proposal.title && !!proposal.problem],
    ["Literature Survey", papers.some(p => p.status === "Reviewed")],
    ["Documentation", chapters.some(c => documentText(docs[c] || "").trim().length > 0)],
    ["Final Report", false],
    ["PPT & Viva", false]
  ];
  const complete = stages.filter(x => x[1]).length;

  return <PageHeader icon="📊" title="Progress Tracker" description="Track your project work as you complete it.">
    <div className="validation-summary"><div><span className="eyebrow">PROJECT PROGRESS</span><h2>{Math.round(complete / stages.length * 100)}% Complete</h2><p>{complete} of {stages.length} stages have saved work.</p></div></div>
    <div className="tracker">{stages.map(([name, done], i) => <div className="tracker-item" key={name}><div className="tracker-number">{done ? "✓" : i + 1}</div><div><strong>{name}</strong><small>{done ? "Work recorded" : "Continue this stage when ready."}</small></div><span>{done ? "Complete" : "In progress"}</span></div>)}</div>
  </PageHeader>;
}

function Dashboard({ user, onLogout }) {
  const [active, setActive] = useState("Dashboard");
  const [proposal, setProposal] = useStoredState("autoscribe-proposal", {
    title: "Smart Classroom Availability System",
    problem: "",
    objectives: "",
    solution: "",
    technologies: "React, Node.js, MongoDB"
  });
  const [savedPrompts, setSavedPrompts] = useStoredState("autoscribe-saved-prompts", []);
  const [docs, setDocs] = useStoredState("autoscribe-docs", starterText);
  const [papers, setPapers] = useStoredState("autoscribe-papers", initialPapers);

  function useIdea(idea) {
    setProposal({
      ...proposal,
      title: idea[0],
      problem: idea[1],
      objectives: `Develop ${idea[0]} with useful, testable features.`,
      solution: idea[1],
      technologies: idea[3].join(", ")
    });
    setActive("Proposal & Planning");
  }

  const menu = [
    ["Dashboard", "⌂"],
    ["Project Ideas", "💡"],
    ["Proposal & Planning", "📋"],
    ["Literature Survey", "📚"],
    ["Documentation", "📝"],
    ["Report Builder", "📄"],
    ["Validation", "✓"],
    ["PPT & Viva", "🎤"],
    ["Progress Tracker", "📊"]
  ];

  function renderPage() {
    switch (active) {
      case "Dashboard":
        return <DashboardHome setActive={setActive} user={user} />;
      case "Project Ideas":
        return <ProjectIdeas onUseIdea={useIdea} />;
      case "Proposal & Planning":
        return <ProposalPlanning proposal={proposal} setProposal={setProposal} savedPrompts={savedPrompts} setSavedPrompts={setSavedPrompts} />;
      case "Literature Survey":
        return <LiteratureSurvey papers={papers} setPapers={setPapers} />;
      case "Documentation":
        return <Documentation docs={docs} setDocs={setDocs} />;
      case "Report Builder":
        return <ReportBuilder proposal={proposal} docs={docs} />;
      case "Validation":
        return <Validation proposal={proposal} docs={docs} />;
      case "PPT & Viva":
        return <PPTViva proposal={proposal} docs={docs} />;
      case "Progress Tracker":
        return <ProgressTracker proposal={proposal} docs={docs} papers={papers} />;
      default:
        return <DashboardHome setActive={setActive} user={user} />;
    }
  }

  return <div className="app-shell">
    <aside className="sidebar">
      <div className="sidebar-logo"><div className="brand-mark">✦</div><div><strong>AutoScribe</strong><small>Project Assistant</small></div></div>
      <nav><span className="nav-label">WORKSPACE</span>{menu.map(([name, icon]) => <button key={name} className={active === name ? "nav-active" : ""} onClick={() => setActive(name)}><span>{icon}</span>{name}</button>)}</nav>
      <div className="sidebar-user">
        <div className="avatar">{(user?.name || "S").charAt(0).toUpperCase()}</div>
        <div><strong>{user?.name || "Student"}</strong><small>B.Tech CSE</small><button className="secondary-button" style={{ marginTop: 6 }} onClick={onLogout}>Sign out</button></div>
      </div>
    </aside>
    <main className="main">{renderPage()}</main>
  </div>;
}

function App() {
  const [user, setUser] = useState(() => readStore("autoscribe-current-user", null));

  if (!user) return <Login onLogin={setUser} />;

  return <Dashboard user={user} onLogout={() => {
    localStorage.removeItem("autoscribe-current-user");
    setUser(null);
  }} />;
}

export default App;
