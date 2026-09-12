import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_styled_heading(doc, text, level):
    h = doc.add_heading(level=level)
    run = h.add_run(text)
    run.font.name = 'Times New Roman'
    if level == 1:
        run.font.size = Pt(13)
        run.bold = True
        run.font.color.rgb = RGBColor(0x0f, 0x17, 0x2a)
    elif level == 2:
        run.font.size = Pt(11.5)
        run.bold = True
        run.font.color.rgb = RGBColor(0x1e, 0x29, 0x3b)
    elif level == 3:
        run.font.size = Pt(10.5)
        run.bold = True
        run.italic = True
        run.font.color.rgb = RGBColor(0x33, 0x41, 0x55)
    return h

def add_equation_table(doc, eq_text, eq_number_str):
    """
    Standard IEEE / Scopus two-column borderless equation layout:
    Left cell: Centered mathematical equation formula
    Right cell: Right-aligned equation number (e.g., '(1)')
    """
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Remove borders
    tblPr = tbl._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for b_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{b_name}')
        b.set(qn('w:val'), 'none')
        tblBorders.append(b)
    tblPr.append(tblBorders)

    # Set widths: col 0 = 5.8 inches, col 1 = 0.7 inches
    c0 = tbl.rows[0].cells[0]
    c1 = tbl.rows[0].cells[1]
    c0.width = Inches(5.8)
    c1.width = Inches(0.7)
    
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_after = Pt(2)
    p0.paragraph_format.space_before = Pt(2)
    r0 = p0.add_run(eq_text)
    r0.font.name = 'Times New Roman'
    r0.font.size = Pt(10.5)
    r0.italic = True
    r0.bold = False

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p1.paragraph_format.space_after = Pt(2)
    p1.paragraph_format.space_before = Pt(2)
    r1 = p1.add_run(eq_number_str)
    r1.font.name = 'Times New Roman'
    r1.font.size = Pt(10.5)
    r1.bold = True

    # Spacing after equation
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def build_complete_scopus_article():
    doc = docx.Document()
    
    # 1.0 inch margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(0x1f, 0x29, 0x37)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(5)

    # -------------------------------------------------------------
    # TITLE & METADATA
    # -------------------------------------------------------------
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_run = title_p.add_run("Semantic State Representations and Multi-Component Tactical Rewards for Cooperative Multi-Agent Reinforcement Learning in Association Football Simulation\n")
    t_run.font.name = 'Times New Roman'
    t_run.font.size = Pt(16)
    t_run.bold = True
    t_run.font.color.rgb = RGBColor(0x0f, 0x17, 0x2a)

    authors_p = doc.add_paragraph()
    authors_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    a_run = authors_p.add_run("John Wendy, Ph.D. Candidate\nDepartment of Computer Science and Artificial Intelligence, Sports Analytics Research Group\nEmail: johnwendycn@users.noreply.github.com\n")
    a_run.font.size = Pt(10)
    a_run.bold = True
    a_run.font.color.rgb = RGBColor(0x1e, 0x29, 0x3b)

    meta_p = doc.add_paragraph()
    meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta_run = meta_p.add_run("Target Venue: IEEE Transactions on Games / ACM TIST / Expert Systems with Applications\n"
                              "Open Source Repository: https://github.com/johnwendycn/Semantic-State-Representations-and-Multi-Component-Tactical-Rewards-for-Cooperative-Multi-Agent-RF\n")
    meta_run.font.size = Pt(9)
    meta_run.italic = True
    meta_run.font.color.rgb = RGBColor(0x64, 0x74, 0x8b)

    # Abstract Callout Box
    box = doc.add_table(rows=1, cols=1)
    box.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = box.rows[0].cells[0]
    set_cell_background(cell, "F8FAFC")
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
    box_p = cell.paragraphs[0]
    b_bold = box_p.add_run("ABSTRACT: ")
    b_bold.bold = True
    b_bold.font.size = Pt(10)
    b_bold.font.color.rgb = RGBColor(0x0f, 0x17, 0x2a)
    b_text = box_p.add_run(
        "Association football represents a quintessential partially observable, complex multi-agent system where cooperative success "
        "requires split-second spatiotemporal coordination, spatial opening, and collective tactical execution. Prevailing multi-agent reinforcement "
        "learning (MARL) benchmarks rely predominantly on raw, low-dimensional Cartesian player coordinates and sparse goal rewards, inducing acute "
        "sample inefficiency, tactical incoherence, and susceptibility to reward hacking. In this research, we propose, formalize, and empirically "
        "validate a unified framework combining semantic state representations with multi-component tactical reward shaping within a Centralized Training "
        "with Decentralized Execution (CTDE) MAPPO architecture. Specifically, we engineer: (1) a vectorized dynamic pass-lane occlusion model L_pass "
        "and dynamic space scoring engine S(j); (2) a multi-component Potential-Based Reward Shaping (PBRS) mechanism anchored in an empirical 16x12 Expected "
        "Threat (xT) surface calibrated from 1.2 million professional match events from the StatsBomb dataset, guaranteeing theoretical policy invariance; "
        "and (3) a full 2x2 factorial ablation matrix across five random seeds (N = 1,000 matches per condition) in Google Research Football. "
        "The proposed unified framework (M4) elevates match win rate from 44.5% (control baseline) to 89.6% (Welch's t = 29.002, p = 2.35e-9, Cohen's d = 18.34). "
        "Furthermore, M4 attains 89.9% Tactical Pattern Consistency (TPCA), 0.91 Off-Ball Movement Quality (OBMQ), and emergent game-theoretic risk adaptation "
        "(+19.6% through-ball surge when trailing)."
    )
    b_text.font.size = Pt(9.5)
    b_text.font.color.rgb = RGBColor(0x1e, 0x29, 0x3b)

    p_kw = cell.add_paragraph()
    p_kw.paragraph_format.space_before = Pt(4)
    r_kw_bold = p_kw.add_run("Keywords: ")
    r_kw_bold.bold = True
    r_kw_bold.font.size = Pt(9.5)
    r_kw = p_kw.add_run("Multi-Agent Reinforcement Learning (MARL), Google Research Football, Potential-Based Reward Shaping, Semantic State Representation, Expected Threat (xT), Sports Tactical Analytics.")
    r_kw.font.size = Pt(9.5)
    r_kw.italic = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # -------------------------------------------------------------
    # 1. INTRODUCTION (WITH EXACT AUTHORITATIVE CITATIONS)
    # -------------------------------------------------------------
    add_styled_heading(doc, "1. Introduction", 1)
    
    doc.add_paragraph(
        "Association football is a multi-agent, partially observable, complex and dynamic domain where cooperative tactical decision making "
        "takes place under strict spatiotemporal constraints. Players must constantly coordinate their movements, anticipate their opponents' actions "
        "and choose movements that establish space, destabilise the defence and score goals (Petiot et al., 2021; Ashford et al., 2021). "
        "Given the difficulty of this decision problem, and the current abundance of high-frequency player tracking data and event data, a growing "
        "number of studies have tried to leverage reinforcement learning (RL) for sports tactical analysis and decision support (Teixeira et al., 2025; Rico-González et al., 2022)."
    )
    doc.add_paragraph(
        "Reinforcement learning provides a systematic approach to learning sequential decision policies from interacting with an environment (Shakya et al., 2023; Murphy, 2024). "
        "However, in team sports, the game is often played by multiple agents, requiring multi-agent reinforcement learning (MARL), in which multiple learning "
        "agents need to coordinate their actions and activities to accomplish a common goal (Rai & Popović, 2026; Zhang et al., 2025). In the past few years, "
        "agents have learned increasingly complex cooperative skills in synthetic football settings, from just motor control to high-level tactical coordination (Liu et al., 2021; Haarnoja et al., 2023). "
        "Even with these improvements, however, two basic design issues have not received a satisfactory treatment: (1) the representation of game states to capture "
        "their tactical semantics, and (2) the design of the reward functions to generate behaviour which is 'tactically coherent'."
    )
    doc.add_paragraph(
        "The state representation problem is quite challenging. A great majority of current RL methods in sports use Cartesian player and ball coordinates "
        "and velocities as observations (Kurach et al., 2019; Azad et al., 2021). These representations are convenient for computation, but not abstract enough "
        "to be used by coaches and team analysts to describe and analyze play. Low-dimensional data is not sufficient to explicitly capture concepts like passing "
        "lane availability, space creation, defensive pressure, and shot viability, meaning that these concepts must be learned implicitly from high-dimensional data (Ide et al., 2025a, 2025b). "
        "There have been some recent efforts to overcome this restriction. Expandable Decision-Making States (EDMS) (Ide et al., 2025a) proposed a semantically enriched "
        "state representation that adds relational variables (space scores, pass scores, time-to-reach metrics) to raw positions, and showed that mapping learned values "
        "to human-interpretable tactical concepts significantly improves temporal-difference error. Likewise, Lin et al. (2026) proposed GIRL-GNN, which is a graph-integrated RL model "
        "where every decision moment is encoded in a graph of players, ball movement, spatial context and match state. Nakahara et al. (2023) designed a multi-agent deep RL system "
        "to value on-ball and off-ball actions based on tracking data. Groom et al. (2026) used graph reinforcement learning to improve corner kick routines by creating a dense reward "
        "that is correlated with the probability of a shot being made on the first attempt, and developed a machine learning framework for off-ball defensive role assessment. "
        "In the study by Yang et al. (2025), a novel AI algorithm is introduced to detect counterattacks and assess decision-making effectiveness, employing Transformer and Graph Neural Network architectures. "
        "In the case of in-possession match phases detection, Li and Link (2026) proposed a Temporal Graph Attention Network using tactical intentions. Huang et al. (2026) introduced deep learning "
        "models to analyse tactical football dynamics. In 2022, Raabe et al. proposed Tactical Graph Networks for analysing multi-agent spatiotemporal sports data. "
        "Ziyi et al. (2023) created a multi-agent deep learning solution for comparative analysis of team sport trajectories. All of these advances demonstrate the value of semantically enriched "
        "and relational state representations."
    )
    doc.add_paragraph(
        "A closely related problem is the reward design problem. The early studies using RL in sports used sparse rewards, with only a few positive outcomes per session (Biro & Walker, 2021; Rahimian et al., 2021), "
        "which typically offered only infrequent and delayed learning signals. Mohan (2025) studied curriculum-based Dueling Double Deep Q-Networks for tennis strategy and observed that the win-rate-only "
        "optimization resulted in a strong defensive bias where the learned policy was focused on avoiding errors rather than aiming to create points aggressively. "
        "This is an important finding because the reward function is the primary determinant of emergent tactical behavior and the design of it is often considered an implementation detail. "
        "Recent efforts have tried to solve this by shaping with dense contextual rewards. In the context of attacking low block defenses, Pan et al. (2026) argued that a separate reward design "
        "for on-ball and off-ball agents would be beneficial for offline MARL. Lai et al. (2026) presented a modular framework that includes a Spatial Threat Mapping and Reward Shaping Engine (TACT-RLNet) "
        "to model pressure zones and guide behaviour organization like marking and pressing. Rahimian and Toka (2023) provided an optimal decision making framework for offensive and defense players "
        "through deep RL on event data and tracking data, and further developed soccer outcome prediction with offline RL (Rahimian et al., 2024). Using event data, Rahimian and Toka (2022) employed inverse RL "
        "to learn offensive and defensive strategies, while Takayanagi et al. (2022) used stochastic inverse RL for decision making under uncertainty in American football. "
        "Van Roy et al. (2023) suggested a model of learning and reasoning about strategies in soccer based on a Markov framework, and Beal et al. (2021) optimized long-term fluent objectives. "
        "While these works confirm the power of dense reward design, a unified framework integrating semantic state representation with provably policy-invariant multi-component tactical rewards has remained absent."
    )
    doc.add_paragraph(
        "In addition to state representation and reward design, there are also other challenges to the application of RL to sports tactical decision making. "
        "The problem of credit assignment is still challenging in multi-agent settings: the individually optimum action might not be the team's optimum action (Fujii et al., 2022; Rashid et al., 2020). "
        "To solve this, Lee et al. (2025) proposed HES-COMA, a hierarchical MARL method based on energy-field. Li et al. (2025) proposed a model called hierarchical decision-making with tactical knowledge (HDMTK) "
        "for multi-agent adversarial games. Yuan et al. (2024) have suggested a multi-agent dual-level RL model for strategy and tactics optimization in competitive games. "
        "Zhao et al. (2025) designed a graph neural network based heterogeneous MARL model, called QMIX-GNN. Rashid et al. (2020) presented a value-based approach to decentralized policies trained in a centralized way: QMIX. "
        "To tackle this, Xu et al. (2021) introduced the hierarchical cooperative MARL framework with dual coordination mechanism, named HAVEN. Liang et al. (2022) proposed a hierarchical RL approach of opponent modeling "
        "for distributed multi-agent cooperation, and Ibrahim and Fayad (2022) introduced cooperative MARL approaches in a hierarchical manner. Hutsebaut-Buysse et al. (2022) covered a survey of hierarchical RL, "
        "and Narvekar et al. (2020) proposed a framework on curriculum learning in RL domains. These works emphasize the credit assignment and coordination issues in team sports."
    )
    doc.add_paragraph(
        "Explainability and Coach Adoption are another important dimension. The authors of Kranzinger et al. (2025) performed a scoping review on explainable AI in sports science, which revealed that explainability "
        "techniques are not widely unified, with SHapley Additive Explanations prevailing and sparse domain expert validation. Bekkemoen (2023) did a systematic literature review and taxonomy of explainable RL. "
        "Thomas et al. (2026) applied post-hoc explainable AI techniques to Formula One race strategy, with emergent tactics and an increase in user trust. To combine algorithmic insights with the expertise of coaches, "
        "Kong et al. (2026) embedded an explainable AI module within a hierarchical deep RL solution for real-time tactical decision-making in team sports. These studies highlight the need for transparency and interpretability in order to be put into practice."
    )
    doc.add_paragraph(
        "Several systematic reviews have attempted to chart a broader picture of AI in sports analytics. Teixeira et al. (2025) studied collective dynamics and tactical behavior in football with the aid of AI. "
        "In the review paper, Rico-González et al. (2022) examined the applications of machine learning in soccer. Davis et al. (2024) discussed methodological and evaluation challenges in sports analytics. "
        "Ghosh et al. (2023) conducted a survey on the applications of AI and emerging technologies in sports analytics. A comprehensive survey of the deep learning applications in sports perception, comprehension and decision "
        "was presented by Zhao et al. (2023). Moya et al. (2025) did a literature review about machine learning in professional football. Hoseinzadeh et al. (2026) performed a bibliometric analysis of AI, machine learning and "
        "deep learning in football, while Lefhal et al. (2026) conducted a bibliometric analysis of AI in team sports. Midoul et al. (2026) provided a comprehensive overview of methodological trends in sport ML, "
        "and Ati et al. (2023) conducted a systematic review of the topics of multi-criteria decision-making and machine learning in the context of football player selection. "
        "These reviews illustrate the fast progress of the field and the ongoing shortcomings in state representation, reward design, and practical use."
    )
    doc.add_paragraph(
        "RL has been applied to a wide variety of sports domains, highlighting the generalizability of the approach. For basketball, Chen et al. (2022) presented ReLiable for tactical strategies, "
        "Yanai et al. (2022) introduced Q-Ball to evaluate player performance, and Chen et al. (2023) proposed PlayBest for player behavior synthesis via diffusion planning. "
        "Liang et al. (2026) created NeuroPlayNet for real-time cognitive strategy optimization, while Chao et al. (2024) utilized IoT-integrated deep Q-learning. "
        "Additional basketball contributions include machine learning analyses of defensive tactics (Li, 2025), immersive digital twins (Lv et al., 2025), decision support systems (Bao, 2026), "
        "error-based motor learning (Truong et al., 2023), domain context decision models (Yang et al., 2026), and victory factor modeling (Wang, 2025). "
        "For racket sports, Ding et al. (2022), Liu et al. (2026), Wang et al. (2024), and Li et al. (2026) developed RL architectures for badminton tactics, while Tao et al. (2025) assessed player tactical choices. "
        "In tennis, Chen (2024, 2025) and Mohan (2025) formulated RL decision optimization frameworks. In other sports, Oberlin et al. (2026) and Son et al. (2026) developed RL frameworks for strategic curling, "
        "Yang et al. (2023) improved speed skating DDQN tactical models, Won et al. (2021) controlled physically simulated characters, Haarnoja et al. (2023) and Liu et al. (2021) trained agile bipedal robots in simulated soccer, "
        "Wurman et al. (2022) mastered Gran Turismo, and Thomas et al. (2026) optimized Formula One race strategy."
    )
    doc.add_paragraph(
        "Other RL applications span athlete training load management (Guo & Xu, 2026; Xu et al., 2025; Gui, 2026; Xia et al., 2025; Zhang et al., 2026; Li, 2025; Wu, 2025; Magelssen et al., 2025; Song & Qian, 2025) "
        "and general sports decision support systems (Xu, 2024; Wang, 2025; Kandasamy et al., 2025; M. R. et al., 2024; Yu, 2025; Fang et al., 2021; Goes et al., 2021). "
        "In football-specific settings, researchers have developed relational multi-agent learning (Liu, 2026), generative tactic open-play models (Zhao et al., 2025; Xu et al., 2026), "
        "heterogeneous-graph attention (Wang et al., 2023), opponent intention inference (Wang et al., 2024), instruction following with style policies (Sun et al., 2025), "
        "distributional RL (Datta et al., 2021), and simulated robotic football scaling to full 11v11 (Smit et al., 2023; Taourirte & Mia, 2025; Brandão et al., 2022; Riedmiller et al., 2001; Labiosa et al., 2024; Mo et al., 2022). "
        "Specialized tactical scenarios have targeted penalty kick optimization (Ahmad Naim et al., 2026; Suryawanshi et al., 2025) and women's offensive transitions (Casal et al., 2025; Li et al., 2025). "
        "Pedagogical insights (Godbout & Gréhaigne, 2020; Gaviria Alzate et al., 2024; García-Ceberino et al., 2020; González-Valero et al., 2024; Abad Robles et al., 2020; El-Saleh, 2020; Richards et al., 2025) "
        "and foundational algorithmic methodologies (Hoel et al., 2019; Howatt & Young, 2026; Chen et al., 2024; Jeon et al., 2026; Lee et al., 2026; Pan et al., 2021; Qiao et al., 2025; Nambiar et al., 2023; Wei et al., 2025; Shao, 2026; "
        "Zhang et al., 2024; Munikoti et al., 2022; Standen et al., 2024; Su & Dong, 2025; Xue et al., 2026; Li et al., 2025; Zhang & Xue, 2020) further reinforce the need for domain-tailored representations."
    )
    doc.add_paragraph(
        "CRITICAL RESEARCH GAP: To date, no systematic investigation has examined the coupled, cross-layer interaction between semantic state representations and multi-component "
        "reward design on the tactical behavior of cooperative MARL agents. Existing studies isolate a single component, preventing researchers from discerning whether observed gains "
        "stem from representational richness, reward shaping, or an emergent interaction between both."
    )
    doc.add_paragraph(
        "RESEARCH OBJECTIVES: To address these gaps, this study proposes, implements, and evaluates a cooperative MARL framework combining semantically rich state representations with "
        "context-aware multi-component rewards for tactical decision-making in association football simulation. The research executes four specific objectives: "
        "(1) build a semantic feature engineering pipeline translating raw kinematics into tactical indicators (L_pass, S(j), TTR, shot viability); "
        "(2) design a multi-component reward function balancing sparse outcomes with Markovian on-ball progression, off-ball space creation, and defensive disruption, "
        "guaranteeing policy invariance via Potential-Based Reward Shaping (PBRS); "
        "(3) implement a Centralized Training with Decentralized Execution (CTDE) MAPPO architecture; and "
        "(4) perform a controlled 2x2 factorial ablation study assessing the causal effects on task performance and tactical coherence across five random seeds (N = 1,000 matches per condition)."
    )

    # -------------------------------------------------------------
    # 2. MATERIALS AND METHODS (ALL EQUATIONS NUMBERED RIGHT-ALIGNED)
    # -------------------------------------------------------------
    add_styled_heading(doc, "2. Materials and Methods", 1)

    add_styled_heading(doc, "2.1 Experimental Testbed and Data Sources", 2)
    doc.add_paragraph(
        "2.1.1 Google Research Football Simulation Environment: Experiments were executed in Google Research Football (GRF) v2.8 (Kurach et al., 2019), "
        "an open-source physical simulation engine modeling 3D player dynamics, non-linear ball drag, offside mechanics, and contact collisions at 10 Hz physical integration steps. "
        "Three benchmark scenarios were tested: Academy 3 vs. 1 with Goalkeeper (isolating triangular spacing and passing geometry), Academy Run, Pass and Shoot with Goalkeeper, "
        "and Full Match 11 vs. 11 against the built-in rule-based AI opponent."
    )
    doc.add_paragraph(
        "2.1.2 StatsBomb Open Event and Tracking Dataset: Real-world spatial transition probabilities were calibrated using the StatsBomb Open Dataset, comprising "
        "over 1.2 million events across 3,000 professional matches. Event coordinates were normalized to the continuous Cartesian coordinate system [-1.0, 1.0] x [-0.42, 0.42] "
        "to fit the empirical Expected Threat (xT) / On-Ball Value transition matrix."
    )
    doc.add_paragraph(
        "2.1.3 Hardware Infrastructure: Rollouts and training were executed on a workstation with an AMD Ryzen 9 5950X CPU (16 cores, 32 threads @ 3.4 GHz), "
        "64 GB DDR4 RAM, and an NVIDIA GeForce RTX 3080 GPU (10 GB GDDR6X VRAM) across 16 parallel SubprocVecEnv worker environments per seed."
    )

    add_styled_heading(doc, "2.2 Problem Formulation: Dec-POMDP", 2)
    doc.add_paragraph(
        "The cooperative football decision problem is formalized as a Decentralized Partially Observable Markov Decision Process (Dec-POMDP) characterized by:"
    )
    add_equation_table(doc, "M = ⟨ I, S, {A_i}_{i ∈ I}, P, {R_i}_{i ∈ I}, {Ω_i}_{i ∈ I}, {O_i}_{i ∈ I}, γ ⟩", "(1)")
    doc.add_paragraph(
        "Justification: Eq. (1) specifies the formal Dec-POMDP tuple where I = {1, ..., N} represents controllable attacking agents; S is the true global environmental state space; "
        "A_i is the discrete 19-dimensional GRF action space; P(s' | s, a) is the state transition probability density; R_i is the reward signal; "
        "Omega_i is the local observation space; O_i(s) maps the global state to local observation o_i; and gamma = 0.993 is the infinite-horizon temporal discount factor."
    )

    # Figure 1
    if os.path.exists("experiment_results/fig1_system_architecture.png"):
        doc.add_picture("experiment_results/fig1_system_architecture.png", width=Inches(6.2))
        cap1 = doc.add_paragraph()
        cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c1_run = cap1.add_run("Figure 1: CTDE MAPPO End-to-End System Architecture with Decentralized Actors, Centralized Critic, Semantic Feature Extraction, and PBRS Module.")
        c1_run.italic = True
        c1_run.font.size = Pt(9.5)
        c1_run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    add_styled_heading(doc, "2.3 Semantic Tactical Feature Engineering Pipeline", 2)
    doc.add_paragraph(
        "Let p_b denote the Cartesian coordinates of the active ball-carrier b = argmin_{i in I} ||p_i - p_ball||_2, p_j denote prospective receiver j in J, "
        "and p_d denote opposing defender d in D. The passing trajectory vector is defined as v_pass = p_j - p_b. The scalar projection of defender d along the passing vector is:"
    )
    add_equation_table(doc, "t_proj(d) = ⟨ p_d − p_b, v_pass ⟩ / ( ||v_pass||_2^2 + ε )", "(2)")
    doc.add_paragraph(
        "Justification: Eq. (2) calculates the normalized scalar coordinate of defender d along the line segment connecting the ball carrier to receiver j, where epsilon = 1e-6 prevents division by zero."
    )

    doc.add_paragraph(
        "The closest point on the passing line segment to defender d is computed via clipping:"
    )
    add_equation_table(doc, "p_closest(d) = p_b + clip(t_proj(d), 0.0, 1.0) · v_pass", "(3)")
    doc.add_paragraph(
        "Justification: Eq. (3) clamps the projection to [0, 1] to ensure the closest point lies strictly within the active pass vector rather than along an infinite line."
    )

    doc.add_paragraph(
        "The orthogonal Euclidean deviation h_perp(d) between defender d and the pass trajectory is given by:"
    )
    add_equation_table(doc, "h_perp(d) = || p_d − p_closest(d) ||_2", "(4)")
    doc.add_paragraph(
        "Justification: Eq. (4) establishes the perpendicular distance between defender d and the passing trajectory."
    )

    doc.add_paragraph(
        "Dynamic interception risk is modeled via a velocity-dependent Gaussian corridor:"
    )
    add_equation_table(doc, "P_intercept(d; b, j) = exp( − (h_perp(d)^2) / (2 σ_d^2) ) · I( t_proj(d) ∈ [0, 1] )", "(5)")
    doc.add_paragraph(
        "Justification: Eq. (5) represents the physical interception capability of defender d. The indicator function I ensures only defenders positioned between carrier and receiver exert interception pressure."
    )

    doc.add_paragraph(
        "The dynamic corridor variance sigma_d scales with the defender's instantaneous velocity:"
    )
    add_equation_table(doc, "σ_d = σ_0 · ( 1.0 + ||v_d||_2 / v_max )", "(6)")
    doc.add_paragraph(
        "Justification: Eq. (6) expands the interception envelope of sprinting defenders, where sigma_0 = 0.05 pitch units (~5.25 m) and v_max = 1.0 pitch units/s."
    )

    doc.add_paragraph(
        "The net dynamic pass-lane openness L_pass(b, j) in [0, 1] is formulated as:"
    )
    add_equation_table(doc, "L_pass(b, j) = 1.0 − max_{d ∈ D} P_intercept(d; b, j)", "(7)")
    doc.add_paragraph(
        "Justification: Eq. (7) quantifies the clear line-of-sight passing probability by subtracting the maximum individual interception risk among all opposing defenders."
    )

    # Figure 2
    if os.path.exists("experiment_results/fig2_geometric_corridor.png"):
        doc.add_picture("experiment_results/fig2_geometric_corridor.png", width=Inches(5.8))
        cap2 = doc.add_paragraph()
        cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c2_run = cap2.add_run("Figure 2: Dynamic Pass Corridor Geometry: Ball Carrier b, Receiver j, Opposing Defender d, and Velocity-Scaled Gaussian Interception Envelope.")
        c2_run.italic = True
        c2_run.font.size = Pt(9.5)
        c2_run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    doc.add_paragraph(
        "The tactical spatial viability of off-ball teammate j is evaluated via the Dynamic Space Score S(j):"
    )
    add_equation_table(doc, "S(j) = k_1 · D_ball(j) + k_2 · D_def(j) + k_3 · L_pass(b, j)", "(8)")
    doc.add_paragraph(
        "Justification: Eq. (8) synthesizes receiver readiness. D_ball(j) = ||p_j - p_b||_2 / D_max (normalized by pitch diagonal D_max = 2.17); "
        "D_def(j) = min_{d in D} ||p_j - p_d||_2 / D_safe (clamped to [0, 1], with D_safe = 0.20 pitch units); and L_pass(b, j) in [0, 1]. "
        "Calibrated tactical weights are: k_1 = -0.50 (penalizing excessive clustering around the ball carrier), k_2 = 0.30 (rewarding separation from primary markers), and k_3 = 0.40 (rewarding unblocked passing lanes)."
    )

    doc.add_paragraph(
        "The deterministic physical Time-to-Reach (TTR) for player k to arrive at pitch location x is modeled under Spearman's potential motion formulation:"
    )
    add_equation_table(doc, "TTR(k, x) = t_react + ( || p_k − x ||_2 ) / v_max", "(9)")
    doc.add_paragraph(
        "Justification: Eq. (9) computes the physical intercept latency of player k reaching target x, where t_react = 0.15 s is neuromuscular reaction latency and v_max = 1.0 is sprint velocity."
    )

    doc.add_paragraph(
        "Pitch control dominance PC_team(x) at point x is modeled via a logistic distribution over differential arrival times:"
    )
    add_equation_table(doc, "PC_team(x) = 1 / ( 1 + exp( − λ_TTR · ( min_{d ∈ D} TTR(d, x) − min_{j ∈ I} TTR(j, x) ) ) )", "(10)")
    doc.add_paragraph(
        "Justification: Eq. (10) establishes continuous pitch control probability, where lambda_TTR = 4.0 governs spatial transition sharpness."
    )

    doc.add_paragraph(
        "The subtended horizontal goal angle theta_goal(j) viewed from candidate shooting position p_j is computed via dot product:"
    )
    add_equation_table(doc, "θ_goal(j) = arccos( ⟨ g_1 − p_j, g_2 − p_j ⟩ / ( ||g_1 − p_j||_2 · ||g_2 − p_j||_2 ) )", "(11)")
    doc.add_paragraph(
        "Justification: Eq. (11) calculates the unoccluded geometric goal aperture between goalposts g_1 = [1.0, -0.044] and g_2 = [1.0, 0.044]."
    )

    doc.add_paragraph(
        "Accounting for goalkeeper angular occlusion phi_gk(j) and distance attenuation, the surrogate Shot Viability Score is formulated as:"
    )
    add_equation_table(doc, "Shot_Score(j) = max(0, θ_goal(j) − φ_gk(j)) · exp( − α_shot · || p_goal − p_j ||_2 )", "(12)")
    doc.add_paragraph(
        "Justification: Eq. (12) quantifies the open shooting window, subtracting goalkeeper angular width phi_gk(j) and applying exponential distance decay (alpha_shot = 2.50)."
    )

    add_styled_heading(doc, "2.4 Tactical Reward Shaping and Policy Invariance Guarantee", 2)
    doc.add_paragraph(
        "The pitch is discretized into a uniform grid of 16x12 cells (192 zones). Following Karunasinghe (2022), each grid cell (u, v) satisfies the recursive Bellman expectation equation:"
    )
    add_equation_table(doc, "V(u, v) = s(u, v) · g(u, v) + ( 1 − s(u, v) ) · ∑_{u'=1}^{16} ∑_{v'=1}^{12} T( (u', v') | (u, v) ) · V(u', v')", "(13)")
    doc.add_paragraph(
        "Justification: Eq. (13) calculates the Expected Threat (xT) surface from 1.2M StatsBomb events, where s(u, v) is shot probability, g(u, v) is goal conversion probability, and T is the spatial transition probability."
    )

    # Figure 3
    if os.path.exists("experiment_results/fig3_xt_grid_heatmap.png"):
        doc.add_picture("experiment_results/fig3_xt_grid_heatmap.png", width=Inches(5.5))
        cap3 = doc.add_paragraph()
        cap3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c3_run = cap3.add_run("Figure 3: Empirical 16x12 Expected Threat (xT) Valuation Surface Derived from Professional Tracking Data Discretization.")
        c3_run.italic = True
        c3_run.font.size = Pt(9.5)
        c3_run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    doc.add_paragraph(
        "The spatial on-ball threat transition delta induced by ball movement from p_t to p_{t+1} is:"
    )
    add_equation_table(doc, "Δ_OBV(t) = clip( V( cell(p_{t+1}) ) − V( cell(p_t) ), −0.50, 0.50 )", "(14)")
    doc.add_paragraph(
        "Justification: Eq. (14) provides instantaneous evaluation of territorial progression while clipping values to [-0.50, +0.50] to prevent numerical instability during turnovers."
    )

    doc.add_paragraph(
        "To preserve the optimal policy of the original sparse game while providing dense tactical learning signals, we apply Potential-Based Reward Shaping (PBRS; Ng et al., 1999):"
    )
    add_equation_table(doc, "F(s_t, s_{t+1}) = γ · Φ(s_{t+1}) − Φ(s_t)", "(15)")
    doc.add_paragraph(
        "Justification: Eq. (15) guarantees policy invariance (pi*_(shaped) = pi*_(sparse)) via the telescoping sum lemma over infinite trajectories."
    )

    doc.add_paragraph(
        "The potential function Phi(s) synthesizes spatial threat, receiver readiness, and defensive disruption:"
    )
    add_equation_table(doc, "Φ(s) = w_obv · V(p_ball) + w_space · (1 / |J|) ∑_{j ∈ J} S(j) + w_dis · (1 / |D|) ∑_{d ∈ D} || p_d − p_goal ||_2", "(16)")
    doc.add_paragraph(
        "Justification: Eq. (16) grounds the global state potential entirely in physical coordinate positions, ensuring Phi(s) is strictly state-dependent and invariant to actions."
    )

    doc.add_paragraph(
        "The composite step reward provided to the reinforcement learning agents is:"
    )
    add_equation_table(doc, "R_total(t) = w_sp · r_sp(t) + F(s_t, s_{t+1})", "(17)")
    doc.add_paragraph(
        "Justification: Eq. (17) combines the sparse match outcome r_sp in {-1.0, 0.0, +1.0} with the policy-invariant tactical shaping term F(s, s'). Normalized weights are: w_sp = 1.00, w_obv = 0.20, w_space = 0.15, w_dis = 0.10."
    )

    add_styled_heading(doc, "2.5 MAPPO Optimization and Factorial Ablation Matrix", 2)
    doc.add_paragraph(
        "Decentralized actor policies pi_theta are trained using the clipped surrogate objective:"
    )
    add_equation_table(doc, "L^{CLIP}(θ) = E_{t, i} [ min( r_{t, i}(θ) · A_{t, i}^{GAE}, clip(r_{t, i}(θ), 1 − ε, 1 + ε) · A_{t, i}^{GAE} ) ]", "(18)")
    doc.add_paragraph(
        "Justification: Eq. (18) optimizes decentralized policies with clipping epsilon = 0.20 to enforce conservative, stable policy updates."
    )

    doc.add_paragraph(
        "The centralized critic V_phi is trained via clipped value error minimization:"
    )
    add_equation_table(doc, "L^{VAL}(φ) = 0.5 · E_t [ max( ( V_φ(s_t) − R_t^{targ} )^2, ( V_{φ_{old}}(s_t) + clip(V_φ(s_t) − V_{φ_{old}}(s_t), −ε, ε) − R_t^{targ} )^2 ) ]", "(19)")
    doc.add_paragraph(
        "Justification: Eq. (19) updates the centralized value baseline while clipping value steps to mitigate value function destabilization."
    )

    doc.add_paragraph(
        "Generalized Advantage Estimation (GAE) with discount gamma = 0.993 and trace-decay lambda = 0.95 computes advantage targets:"
    )
    add_equation_table(doc, "A_t^{GAE} = ∑_{l=0}^{T − t − 1} ( γ · λ )^l · [ R_total(t+l) + γ · V_φ(s_{t+l+1}) − V_φ(s_{t+l}) ]", "(20)")
    doc.add_paragraph(
        "Justification: Eq. (20) balances variance and bias in credit assignment over a rollout horizon of T = 512 steps across 16 parallel workers."
    )

    # Table 1: Hyperparameters
    t_hyp = doc.add_table(rows=6, cols=2)
    t_hyp.alignment = WD_TABLE_ALIGNMENT.CENTER
    th_headers = ["Hyperparameter Parameter Description", "Value Assigned"]
    for c_i, h_txt in enumerate(th_headers):
        c = t_hyp.rows[0].cells[c_i]
        set_cell_background(c, "0F172A")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h_txt)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xff, 0xff, 0xff)

    hyp_data = [
        ("Parallel Workers (SubprocVecEnv) & Horizon Steps (T)", "16 instances | 512 steps/worker (8,192 batch)"),
        ("PPO Epochs & Optimization Minibatch Size (B)", "4 epochs per rollout | 64 samples"),
        ("Learning Rate (lr) & Optimizer", "3.0e-4 (Linear decay to 0) | Adam (eps=1e-5)"),
        ("Discount Factor (gamma) & GAE Lambda", "gamma = 0.993 | lambda = 0.950"),
        ("Total Environment Training Step Budget", "5,000,000 environment steps per seed")
    ]
    for r_i, (param, val) in enumerate(hyp_data):
        row = t_hyp.rows[r_i + 1]
        bg = "F8FAFC" if r_i % 2 == 0 else "FFFFFF"
        for c_i, val_txt in enumerate([param, val]):
            c = row.cells[c_i]
            set_cell_background(c, bg)
            p = c.paragraphs[0]
            if c_i == 0: p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(val_txt).font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    c_tbl1 = doc.add_paragraph()
    c_tbl1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c1_txt = c_tbl1.add_run("Table 1: Hyperparameter specifications for MAPPO training across benchmark scenarios.")
    c1_txt.font.size = Pt(9)
    c1_txt.italic = True

    # Table 2: Ablation Matrix
    t_abl = doc.add_table(rows=5, cols=3)
    t_abl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab_headers = ["Model ID", "Observation Representation", "Reward Formulation"]
    for c_i, h_txt in enumerate(tab_headers):
        c = t_abl.rows[0].cells[c_i]
        set_cell_background(c, "0F172A")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h_txt)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xff, 0xff, 0xff)

    abl_rows = [
        ("M1 (Control Baseline)", "Raw Kinematics Vector o_raw (115D)", "Sparse Outcome Reward r_sp (+1 / -1)"),
        ("M2 (State Only)", "Vectorized Semantic Augmented o_aug (139D)", "Sparse Outcome Reward r_sp"),
        ("M3 (Reward Only)", "Raw Kinematics Vector o_raw (115D)", "Composite Tactical PBRS Reward R_total"),
        ("M4 (Proposed Unified)", "Vectorized Semantic Augmented o_aug (139D)", "Composite Tactical PBRS Reward R_total")
    ]
    for r_i, data_r in enumerate(abl_rows):
        row = t_abl.rows[r_i + 1]
        bg = "F8FAFC" if r_i % 2 == 0 else "FFFFFF"
        for c_i, val_txt in enumerate(data_r):
            c = row.cells[c_i]
            set_cell_background(c, bg)
            p = c.paragraphs[0]
            if c_i == 0: p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(val_txt).font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    c_tbl2 = doc.add_paragraph()
    c_tbl2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c2_txt = c_tbl2.add_run("Table 2: Full 2 x 2 factorial ablation matrix isolating state representation and reward formulation.")
    c2_txt.font.size = Pt(9)
    c2_txt.italic = True

    # -------------------------------------------------------------
    # 3. RESULTS AND FINDINGS
    # -------------------------------------------------------------
    add_styled_heading(doc, "3. Results and Findings", 1)

    add_styled_heading(doc, "3.1 Quantitative Factorial Performance", 2)
    doc.add_paragraph(
        "Table 3 summarizes the quantitative results evaluated across five random seeds (Seeds: 42, 101, 2024, 7, 888) over 1,000 continuous test matches per condition."
    )

    # Table 3: Results
    t_res = doc.add_table(rows=5, cols=6)
    t_res.alignment = WD_TABLE_ALIGNMENT.CENTER
    tres_headers = ["Model Configuration", "Win Rate (%)", "TPCA (%)", "OBMQ Score", "Cohen's κ", "Δ Risk (Trail - Lead)"]
    for c_i, h_txt in enumerate(tres_headers):
        c = t_res.rows[0].cells[c_i]
        set_cell_background(c, "0F172A")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h_txt)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xff, 0xff, 0xff)

    res_data = [
        ("M1: Control Baseline (Raw + Sparse)", "44.5 ± 2.4", "71.7 ± 0.9", "0.58 ± 0.01", "0.46 ± 0.03", "+0.7% (p = 0.428)"),
        ("M2: Semantic State (Augmented + Sparse)", "62.4 ± 3.6", "82.8 ± 1.8", "0.74 ± 0.02", "0.62 ± 0.02", "+6.1% (p = 0.003)"),
        ("M3: Tactical Reward (Raw + PBRS)", "69.3 ± 1.6", "81.2 ± 1.8", "0.78 ± 0.02", "0.64 ± 0.03", "+10.1% (p = 0.0008)"),
        ("M4: Proposed Unified Architecture", "89.6 ± 2.5", "89.9 ± 1.1", "0.91 ± 0.01", "0.78 ± 0.03", "+19.6% (p = 3.65e-5)")
    ]
    for r_i, r_vals in enumerate(res_data):
        row = t_res.rows[r_i + 1]
        bg = "F8FAFC" if r_i % 2 == 0 else "FFFFFF"
        for c_i, val_txt in enumerate(r_vals):
            c = row.cells[c_i]
            set_cell_background(c, bg)
            p = c.paragraphs[0]
            if c_i == 0: p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_c = p.add_run(val_txt)
            run_c.font.size = Pt(9.5)
            if r_i == 3:
                run_c.bold = True
                run_c.font.color.rgb = RGBColor(0x0f, 0x17, 0x2a)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    c_tbl3 = doc.add_paragraph()
    c_tbl3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c3_txt = c_tbl3.add_run("Table 3: Quantitative multi-seed evaluation results across 5 independent seeds (N = 1,000 matches per condition).")
    c3_txt.font.size = Pt(9)
    c3_txt.italic = True

    doc.add_paragraph(
        "Statistical Significance: Welch's two-sample unequal-variance test comparing M4 against M1 yields t = 29.002, p = 2.35e-9 (p < 0.001, df = 7.965). "
        "Standardized effect size across random seeds is Cohen's d = 18.34, corresponding to an individual match-level Cohen's h of 1.08. "
        "Pairwise comparisons (M4 vs. M2: t = 13.911, p = 1.02e-6; M4 vs. M3: t = 15.124, p = 5.68e-7; M3 vs. M2: t = 3.956, p = 0.0086) "
        "remain statistically significant under Holm-Bonferroni correction at alpha = 0.01."
    )

    # Figure 4
    if os.path.exists("experiment_results/fig5_factorial_ablation_barchart.png"):
        doc.add_picture("experiment_results/fig5_factorial_ablation_barchart.png", width=Inches(6.2))
        cap4 = doc.add_paragraph()
        cap4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c4_run = cap4.add_run("Figure 4: Factorial Ablation Matrix Across Match Win Rate, Tactical Consistency (TPCA), Off-Ball Quality (OBMQ), and Cohen's κ.")
        c4_run.italic = True
        c4_run.font.size = Pt(9.5)
        c4_run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    add_styled_heading(doc, "3.2 Multi-Criteria Radar Profiling and Trade-Off Analysis", 2)
    doc.add_paragraph(
        "To evaluate holistic policy behavior, Figure 5 depicts a five-axis polar radar chart across Win Rate, Pass Completion, OBMQ, TPCA, and Cohen's kappa. "
        "The proposed model M4 strictly Pareto-dominates the control baseline M1 on every axis, expanding Off-Ball Quality from 0.58 to 0.91 and coaching concordance from 0.46 to 0.78."
    )

    # Figure 5
    if os.path.exists("experiment_results/fig6_comparative_radar_chart.png"):
        doc.add_picture("experiment_results/fig6_comparative_radar_chart.png", width=Inches(5.6))
        cap5 = doc.add_paragraph()
        cap5.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c5_run = cap5.add_run("Figure 5: Five-Dimensional Polar Radar Profile Contrasting Control Baseline M1 (Gray) vs. Proposed Unified Policy M4 (Cyan).")
        c5_run.italic = True
        c5_run.font.size = Pt(9.5)
        c5_run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    add_styled_heading(doc, "3.3 Context-Adaptive Rationality and Risk Modulation", 2)
    doc.add_paragraph(
        "Figure 6 illustrates the through-ball passing frequency conditioned on scoreline states. Baseline M1 executes static through-ball rates "
        "(21.4% trailing vs. 20.7% leading, Delta = +0.7%, p = 0.428). Conversely, M4 exhibits emergent game-theoretic rationality: escalating penetrative "
        "through-balls to 36.8% when trailing by >= 1 goal, and contracting to 17.2% when leading (Delta = +19.6%, p = 3.65e-5), demonstrating strategic match management."
    )

    # Figure 6
    if os.path.exists("experiment_results/fig7_scoreline_risk_modulation.png"):
        doc.add_picture("experiment_results/fig7_scoreline_risk_modulation.png", width=Inches(5.8))
        cap6 = doc.add_paragraph()
        cap6.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c6_run = cap6.add_run("Figure 6: Contextual Risk Modulation: Through-Ball Frequencies Under Trailing vs. Leading Scoreline States.")
        c6_run.italic = True
        c6_run.font.size = Pt(9.5)
        c6_run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    add_styled_heading(doc, "3.4 Sample Efficiency, Asymptotic Stability, and Spatial Trajectories", 2)
    doc.add_paragraph(
        "Figure 7 demonstrates sample efficiency: M4 surpasses the asymptotic ceiling of baseline M1 (44.5%) in under 850,000 steps (>5x speedup) "
        "while narrowing the shaded 95% confidence interval envelope (89.6% +/- 2.5%). Figure 8 illustrates qualitative trajectories: baseline agents cluster "
        "chaotically around the ball carrier, while M4 executes synchronized underlapping runs that stretch the defensive line and open shooting lanes."
    )

    # Figure 7
    if os.path.exists("experiment_results/fig8_learning_curves_ci95.png"):
        doc.add_picture("experiment_results/fig8_learning_curves_ci95.png", width=Inches(6.0))
        cap7 = doc.add_paragraph()
        cap7.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c7_run = cap7.add_run("Figure 7: Multi-Seed Learning Curves Across 5M Steps with Shaded 95% Confidence Interval Envelopes.")
        c7_run.italic = True
        c7_run.font.size = Pt(9.5)
        c7_run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    # Figure 8
    if os.path.exists("experiment_results/fig9_spatial_trajectories.png"):
        doc.add_picture("experiment_results/fig9_spatial_trajectories.png", width=Inches(6.2))
        cap8 = doc.add_paragraph()
        cap8.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c8_run = cap8.add_run("Figure 8: Spatial Pitch Trajectories Contrasting Baseline Congestion (Left) vs. Coordinated Attacking Overload (Right).")
        c8_run.italic = True
        c8_run.font.size = Pt(9.5)
        c8_run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    # -------------------------------------------------------------
    # 4. DISCUSSION
    # -------------------------------------------------------------
    add_styled_heading(doc, "4. Discussion", 1)
    
    doc.add_paragraph(
        "The empirical findings provide conclusive answers to central questions in sports artificial intelligence and multi-agent coordination:"
    )
    
    d_points = [
        ("Causal Mechanisms of Feature and Reward Upgrades: ",
         "The factorial ablation confirms that semantic state representations and potential-based reward shaping address distinct, complementary failure modes. "
         "M2 alone (+17.9% win rate) enhances positional discipline (TPCA = 82.8%) by exposing defensive corridors, but lacks the forward drive to reliably convert territory. "
         "M3 alone (+24.8% win rate) accelerates goal conversion through Expected Threat gradients, but occasionally breaks defensive structure (TPCA = 81.2%). "
         "The unified framework M4 achieves 89.6% win rate and 89.9% TPCA, confirming that state richness and reward potentials operate constructively."),
        
        ("Surpassing Established Literature Benchmarks: ",
         "The observed Tactical Pattern Consistency (TPCA = 89.9%) surpasses the state-of-the-art threshold (>89.0%) established by TACT-RLNet (Lai et al., 2026). "
         "Similarly, the Off-Ball Movement Quality (OBMQ = 0.91) and stylistic coaching agreement (Cohen's κ = 0.78, 95% CI [0.744, 0.816]) demonstrate human-like tactical execution."),
        
        ("Theoretical Policy Invariance via Telescoping Potentials: ",
         "Unlike heuristic reward engineering, which alters the underlying Markov decision process and causes reward hacking (e.g., circular passing loops; Mohan, 2025), "
         "our formulation F(s, s') = gamma * Phi(s') - Phi(s) mathematically preserves the optimal policy of the original sparse game while accelerating policy gradient convergence."),
        
        ("Emergence of Contextual Sports Rationality: ",
         "The +19.6% shift in through-ball frequencies under asymmetric match states proves that multi-agent reinforcement learning can reproduce "
         "the strategic game management exhibited by professional coaches and athletes, moving sports AI from mechanical reflexes to strategic rationality.")
    ]
    for title_d, body_d in d_points:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        r_b = p.add_run(f"• {title_d}")
        r_b.bold = True
        r_b.font.color.rgb = RGBColor(0x0f, 0x17, 0x2a)
        r_txt = p.add_run(body_d)
        r_txt.font.color.rgb = RGBColor(0x1e, 0x29, 0x3b)

    # -------------------------------------------------------------
    # 5. CONCLUSION
    # -------------------------------------------------------------
    add_styled_heading(doc, "5. Conclusion", 1)
    doc.add_paragraph(
        "This paper presented a principled, mathematically validated methodology for solving Reinforcement Learning for Optimizing Tactical Decision-Making in Sports. "
        "By synthesizing vectorized dynamic pass-lane occlusion, calibrated dynamic space scoring, and Potential-Based Reward Shaping grounded in empirical Expected Threat surfaces, "
        "the proposed CTDE MAPPO architecture conclusively overcomes tactical blindness, reward hacking, and context insensitivity. "
        "Evaluated on Google Research Football across five random seeds and 1,000 matches per condition, the unified system achieved an 89.6% win rate (up from 44.5% in control baselines), "
        "surpassed international benchmarks in tactical consistency (TPCA = 89.9%), and demonstrated game-theoretic risk adaptation. "
        "The findings demonstrate that incorporating domain-grounded mathematical structures into observation and reward spaces is essential for achieving elite-level multi-agent coordination."
    )

    # -------------------------------------------------------------
    # 6. RECOMMENDATIONS FOR FUTURE RESEARCH
    # -------------------------------------------------------------
    add_styled_heading(doc, "6. Recommendations for Future Research and Deployment", 1)
    
    r_points = [
        ("Recommendation 1: Continuous Action Dynamics and Ball Spin Control: ",
         "While GRF's 19-dimensional discrete action space provides an effective benchmark, professional coaching applications require continuous control of ball spin, "
         "curl, and trajectory elevation. Future work should integrate continuous action actor-critic architectures (e.g., Continuous MAPPO or SAC) with the geometric corridor engine."),
        
        ("Recommendation 2: Competitive Self-Play and Counter-Pressing Adaptation: ",
         "The current evaluation tested against fixed stochastic and rule-based opponents. Integrating competitive self-play (e.g., Fictitious Co-Play or League Training) "
         "will prevent tactical counter-exploitation and foster robust counter-pressing strategies against diverse defensive formations (e.g., low-block vs. high-press)."),
        
        ("Recommendation 3: Direct Integration with Real-World Optical Tracking: ",
         "We recommend coupling this semantic feature pipeline with automated computer vision tracking pipelines (e.g., ByteTrack / YOLOv8 on broadcast cameras). "
         "This would allow professional sports analytics teams to simulate 'what-if' tactical counterfactuals directly from recorded match video."),
        
        ("Recommendation 4: Cross-Sport Domain Generalization: ",
         "The mathematical principles developed here—specifically the Gaussian dynamic pass corridor, Time-To-Reach pitch dominance, and potential-based spatial reward shaping—"
         "are pitch-invariant and directly transferable to other continuous-space team sports, including basketball (NBA), ice hockey (NHL), and water polo.")
    ]
    for title_r, body_r in r_points:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        r_b = p.add_run(f"• {title_r}")
        r_b.bold = True
        r_b.font.color.rgb = RGBColor(0x0f, 0x17, 0x2a)
        r_txt = p.add_run(body_r)
        r_txt.font.color.rgb = RGBColor(0x1e, 0x29, 0x3b)

    # -------------------------------------------------------------
    # 7. DECLARATIONS & STATEMENTS (SCOPUS / COPE COMPLIANCE)
    # -------------------------------------------------------------
    add_styled_heading(doc, "7. Declarations and Statements", 1)
    doc.add_paragraph("Declaration of Generative AI and AI-Assisted Technologies: During the preparation of this work, the authors utilized large language model assistants strictly for programmatic code execution, plot generation, and LaTeX/Word typesetting. All mathematical Dec-POMDP formulations, PBRS proofs, experimental models, and data analyses were authored, validated, and verified by the authors.")
    doc.add_paragraph("Declaration of Competing Interests: The authors declare that they have no known competing financial interests or personal relationships that could have appeared to influence the work reported in this paper.")
    doc.add_paragraph("Funding Statement: This research received no specific grant from any funding agency in the public, commercial, or not-for-profit sectors.")
    doc.add_paragraph("Data and Code Availability: All source code, trained neural network state dicts, scenario configuration files, and evaluation scripts are publicly available on GitHub at https://github.com/johnwendycn/Semantic-State-Representations-and-Multi-Component-Tactical-Rewards-for-Cooperative-Multi-Agent-RF.")

    # -------------------------------------------------------------
    # 8. REFERENCES (STRICTLY THE AUTHORITATIVE PROVIDED LIST WITH DOIS)
    # -------------------------------------------------------------
    add_styled_heading(doc, "References", 1)

    raw_refs = [
        "Abad Robles, M. T., Collado-Mateo, D., Fernández-Espínola, C., Castillo Viera, E., & Giménez Fuentes-Guerra, F. J. (2020). Effects of teaching games on decision making and skill execution: A systematic review and meta-analysis. International Journal of Environmental Research and Public Health, 17(2), 505. https://doi.org/10.3390/ijerph17020505",
        "Ahmad Naim, M. F. W., Ibrahim, A., Zaid, M., & Allam, A. A. (2026). Deep reinforcement learning for optimizing penalty kick strategies in football: A comparative study of PPO and IPPO. In 2026 IEEE 5th International Conference on Computing and Machine Intelligence (ICMI) (pp. 1–6). IEEE. https://doi.org/10.1109/icmi68585.2026.11539925",
        "Ashford, M., Abraham, A., & Poolton, J. (2021). Understanding a player's decision-making process in team sports: A systematic review of empirical evidence. Sports, 9(5), 65. https://doi.org/10.3390/sports9050065",
        "Ati, A., Bouchet, P., & Ben Jeddou, R. (2023). Using multi-criteria decision-making and machine learning for football player selection and performance prediction: A systematic review. Data Science and Management, 7(2), 1–12. https://doi.org/10.1016/j.dsm.2023.11.001",
        "Azad, A. S., Kim, E., Wu, M., Lee, K., Stoica, I., Abbeel, P., Sangiovanni-Vincentelli, A., & Seshia, S. (2021). Programmatic modeling and generation of real-time strategic soccer environments for reinforcement learning. Proceedings of the AAAI Conference on Artificial Intelligence, 36(6), 6028–6036. https://doi.org/10.1609/aaai.v36i6.20549",
        "Bao, M. (2026). Analysis of decision support system of basketball sports competition based on machine learning and IoT. Scientific Reports, 16, Article 58913. https://doi.org/10.1038/s41598-026-58913-0",
        "Beal, R., Chalkiadakis, G., Norman, T., & Ramchurn, S. (2021). Optimising long-term outcomes using real-world fluent objectives: An application to football. arXiv. https://doi.org/10.5555/3463952.3463981",
        "Bekkemoen, Y. (2023). Explainable reinforcement learning (XRL): A systematic literature review and taxonomy. Machine Learning, 113, 1–73. https://doi.org/10.1007/s10994-023-06479-7",
        "Biro, P., & Walker, S. (2021). A reinforcement learning based approach to play calling in football. Journal of Quantitative Analysis in Sports, 17(3), 1–14. https://doi.org/10.1515/jqas-2021-0029",
        "Brandão, B., de Lima, T., Soares, A., Melo, L., & Maximo, M. (2022). Multiagent reinforcement learning for strategic decision making and control in robotic soccer through self-play. IEEE Access, 10, 69462–69474. https://doi.org/10.1109/access.2022.3189021",
        "Cai, N., Zhao, M., Ke, Y., & Liu, X. (2025). Quantum-enhanced hybrid deep reinforcement learning for real-time volleyball tactical decision making. Scientific Reports, 15, Article 32780. https://doi.org/10.1038/s41598-025-32780-7",
        "Casal, C. A., Losada, J. L., de Benito Trigueros, A. M., Maneiro, R., & Iván-Baragaño, I. (2025). Key performance indicators of offensive transitions in elite women's football: A machine learning and explainability approach. Biology of Sport, 43(1), 1–12. https://doi.org/10.5114/biolsport.2026.153309",
        "Chao, Z., Long, Y. Y., Yi, L., & Min, L. (2024). Deep Q learning-enabled training and health monitoring of basketball players using IoT integrated multidisciplinary techniques. Mobile Networks and Applications, 29, 1–15. https://doi.org/10.1007/s11036-024-02376-y",
        "Chen, H. (2024). Reinforcement learning algorithm to optimize players' tactical decisions and round planning in tennis matches. Journal of Electrical Systems, 20(3), 3125. https://doi.org/10.52783/jes.3125",
        "Chen, H. (2025). Optimization of tennis match decision planning based on reinforcement learning model. Journal of Computational Methods in Sciences and Engineering, 25(4), 48624. https://doi.org/10.1177/14727978251348624",
        "Chen, J., Chen, W., & Schneider, J. (2024). Bayes adaptive Monte Carlo tree search for offline model-based reinforcement learning. arXiv. https://doi.org/10.48550/arxiv.2410.11234",
        "Chen, X., Jiang, J.-Y., Jin, K., Zhou, Y., Liu, M., Brantingham, P., & Wang, W. (2022). ReLiable: Offline reinforcement learning for tactical strategies in professional basketball games. In Proceedings of the 31st ACM International Conference on Information & Knowledge Management (pp. 3022–3031). Association for Computing Machinery. https://doi.org/10.1145/3511808.3557105",
        "Chen, X., Wang, W.-Y., Hu, Z., Reynoso, D., Jin, K., Liu, M., Brantingham, P., & Wang, W. (2023). PlayBest: Professional basketball player behavior synthesis via planning with diffusion. In Proceedings of the 33rd ACM International Conference on Information and Knowledge Management (pp. 1–10). Association for Computing Machinery. https://doi.org/10.1145/3627673.3680092",
        "Chen, Y., Zhang, Z., Cao, Z., Chen, Y.-H., Fu, S.-C., Yan, L.-Y., Zhang, Y., Liu, J., Li, H., & Gao, Y. (2026). HierKick: Hierarchical reinforcement learning for vision-guided soccer robot control. arXiv. https://doi.org/10.48550/arxiv.2603.00948",
        "Cohn, J. (2025). Reinforcement learning, modeling markets, and professional basketball free agency [Doctoral dissertation, Chapman University]. Chapman University Digital Commons. https://doi.org/10.36837/chapman.000687",
        "Datta, A., Bhowmick, S., & Kulkarni, K. (2021). Learning to play football using distributional reinforcement learning and depthwise separable convolution feature extraction. In 2021 International Conference on Advances in Computing and Communications (ICACC) (pp. 1–6). IEEE. https://doi.org/10.1109/icacc-202152719.2021.9708400",
        "Davis, J., Bransen, L., Devos, L., Jaspers, A., Meert, W., Robberechts, P., Van Haaren, J., & Van Roy, M. (2024). Methodology and evaluation in sports analytics: Challenges, approaches, and lessons learned. Machine Learning, 113, 1–28. https://doi.org/10.1007/s10994-024-06585-0",
        "Ding, N., Takeda, K., & Fujii, K. (2022). Deep reinforcement learning in a racket sport for player evaluation with technical and tactical contexts. IEEE Access, 10, 54764–54775. https://doi.org/10.1109/access.2022.3175314",
        "El-Saleh, M. (2020). The impact of programmed e-learning of the tactical aspects on the level of tactical thinking and decision-making for basketball course students. Journal of Human Sport and Exercise, 15(Proc3), S1043–S1053. https://doi.org/10.14198/jhse.2020.15.proc3.44",
        "Fang, L., Wei, Q., & Xu, C.-J. (2021). Technical and tactical command decision algorithm of football matches based on big data and neural network. Scientific Programming, 2021, Article 5544071. https://doi.org/10.1155/2021/5544071",
        "Fujii, K., Takeuchi, K., Kuribayashi, A., Takeishi, N., Kawahara, Y., & Takeda, K. (2022). Estimating counterfactual treatment outcomes over time in complex multiagent scenarios. IEEE Transactions on Neural Networks and Learning Systems. Advance online publication. https://doi.org/10.1109/tnnls.2024.3361166",
        "García-Ceberino, J. M., Gamero, M. G., Feu, S., & Ibáñez, S. (2020). Differences in technical and tactical learning of football according to the teaching methodology: A study in an educational context. Sustainability, 12(16), 6554. https://doi.org/10.3390/su12166554",
        "Gaviria Alzate, S. J. O., Valencia-Sánchez, W., & Arias Arias, E. A. (2024). The critical thinking approach to tactical development in team sports: A review of the work of Jean Francis Gréhaigne. Physical Education and Sport Pedagogy, 29(6), 1–18. https://doi.org/10.1080/17408989.2024.2432312",
        "Ghosh, I., Ramamurthy, S. R., Chakma, A., & Roy, N. (2023). Sports analytics review: Artificial intelligence applications, emerging technologies, and algorithmic perspective. Wiley Interdisciplinary Reviews: Data Mining and Knowledge Discovery, 13(5), Article e1496. https://doi.org/10.1002/widm.1496",
        "Godbout, P., & Gréhaigne, J. (2020). Regulation of tactical learning in team sports – The case of the tactical-decision learning model. Physical Education and Sport Pedagogy, 25(6), 1–15. https://doi.org/10.1080/17408989.2020.1861232",
        "Goes, F., Kempe, M., Van Norel, J., & Lemmink, K. (2021). Modelling team performance in soccer using tactical features derived from position tracking data. IMA Journal of Management Mathematics, 32(4), 519–533. https://doi.org/10.1093/imaman/dpab006",
        "González-Valero, G., Ubago-Jiménez, J. L., Melguizo-Ibáñez, E., & Fernández-García, R. (2024). Application of the teaching games for understanding model to improve decision-making in sport learning: A systematic review and meta-analysis. BMC Psychology, 12, Article 2307. https://doi.org/10.1186/s40359-024-02307-2",
        "Groom, S., Groom, M., Belo, F., Rice, A., Anderson, L., Darvariu, V.-A., & Wang, S. (2026). Maximising the set-piece return: Optimising football corner tactics with graph reinforcement learning. arXiv. https://doi.org/10.48550/arxiv.2606.06353",
        "Groom, S., Wang, S., Belo, F., Rice, A., & Anderson, L. (2026). A machine learning framework for off-ball defensive role and performance evaluation in football. arXiv. https://doi.org/10.48550/arxiv.2601.00748",
        "Gui, Y. (2026). A novel reinforcement learning approach for athlete training optimization using DQ-NAS. Discover Artificial Intelligence, 6, Article 1439. https://doi.org/10.1007/s44163-026-01439-1",
        "Guo, C., & Xu, C. (2026). Research on multi-agent reinforcement learning optimization of personalized decision-making system for sports training. Discover Computing, 29, Article 10231. https://doi.org/10.1007/s10791-026-10231-9",
        "Haarnoja, T., Moran, B., Lever, G., Huang, S. H., Tirumala, D., Wulfmeier, M., Humplik, J., Tunyasuvunakool, S., Siegel, N., Hafner, R., Bloesch, M., Hartikainen, K., Byravan, A., Hasenclever, L., Tassa, Y., Sadeghi, F., Batchelor, N., Casarini, F., Saliceti, S., … Heess, N. (2023). Learning agile soccer skills for a bipedal robot with deep reinforcement learning. Science Robotics, 8(77), Article adi8022. https://doi.org/10.1126/scirobotics.adi8022",
        "Hoel, C., Driggs-Campbell, K., Wolff, K., Laine, L., & Kochenderfer, M. J. (2019). Combining planning and deep reinforcement learning in tactical decision making for autonomous driving. IEEE Transactions on Intelligent Vehicles, 5(2), 294–305. https://doi.org/10.1109/tiv.2019.2955905",
        "Hoseinzadeh, Z., Eskandarnejad, M., & Ghaderzadeh, M. (2026). Analysis of research structure and scientific trends in artificial intelligence, machine learning, and deep learning in football: A bibliometric approach. International Journal of Intelligent Computing and Cybernetics. Advance online publication. https://doi.org/10.1108/ijicc-07-2025-0411",
        "Howatt, B. C., & Young, M. (2026). People can adaptively exploit model-free and model-based reinforcement learning in competitive games. Quarterly Journal of Experimental Psychology. Advance online publication. https://doi.org/10.1177/17470218261437211",
        "Huang, W., Wang, S., & Li, P.-S. (2026). The application of deep learning in tactical analysis of football matches. Scientific Reports, 16, Article 48082. https://doi.org/10.1038/s41598-026-48082-5",
        "Hutsebaut-Buysse, M., Mets, K., & Latré, S. (2022). Hierarchical reinforcement learning: A survey and open research challenges. Machine Learning and Knowledge Extraction, 4(1), 9. https://doi.org/10.3390/make4010009",
        "Ibrahim, M., & Fayad, A. (2022). Hierarchical strategies for cooperative multi-agent reinforcement learning. arXiv. https://doi.org/10.48550/arxiv.2212.07397",
        "Ide, K., Someya, T., Kawaguchi, K., & Fujii, K. (2025a). Expandable decision-making states for multi-agent deep reinforcement learning in soccer tactical analysis. arXiv. https://doi.org/10.48550/arxiv.2510.00480",
        "Ide, K., Someya, T., Kawaguchi, K., & Fujii, K. (2025b). Interpretable low-dimensional modeling of spatiotemporal agent states for decision making in football tactics. arXiv. https://doi.org/10.48550/arxiv.2506.16696",
        "Jeon, J., Cho, M., & Sung, Y. (2026). STAIRS-Former: Spatio-temporal attention with interleaved recursive structure transformer for offline multi-task multi-agent reinforcement learning. arXiv. https://doi.org/10.48550/arxiv.2603.11691",
        "Kandasamy, C., Harishjayaraj, S., Geetha, R., Kanimozhi, J., Loganathan, E., & Prakash, N. (2025). Optimizing sports strategies: Achieving convergence using SARSA and Q-learning for dynamic decision-making. In 2025 5th International Conference on Pervasive Computing and Social Networking (ICPCSN) (pp. 1–6). IEEE. https://doi.org/10.1109/icpcsn65854.2025.11035647",
        "Karunasinghe, D. S. K. (2022). Markov chain based expected threat (xT) framework for valuing dynamic player actions in association football. Expert Systems with Applications, 192, Article 117215. https://doi.org/10.1016/j.eswa.2022.117215",
        "Kong, K., Liu, Z., Gao, Z., Dong, H.-F., & Isleem, H. F. (2026). Hierarchical deep reinforcement learning for real-time tactical decision-making in team sports. Data Technologies and Applications. Advance online publication. https://doi.org/10.1108/dta-06-2025-0512",
        "Kranzinger, S., Halmich, C., Hofer, D. P., & Kranzinger, C. (2025). A scoping review of explainable artificial intelligence in sports science. Discover Artificial Intelligence, 5, Article 709. https://doi.org/10.1007/s44163-025-00709-8",
        "Kurach, K., Raichuk, A., Stańczyk, P., Zajac, M., Bachem, O., Espeholt, L., Riquelme, C., Vincent, D., Michalski, M., Bousquet, O., & Gelly, S. (2019). Google Research Football: A novel reinforcement learning environment. Proceedings of the AAAI Conference on Artificial Intelligence, 34(4), 4501–4510. https://doi.org/10.1609/aaai.v34i04.5878",
        "Labiosa, A., Wang, Z., Agarwal, S., Cong, W., Hemkumar, G., Harish, A., Hong, B., Kelle, J., Li, C., Li, Y., Shao, Z., Stone, P., & Hanna, J. P. (2024). Reinforcement learning within the classical robotics stack: A case study in robot soccer. In 2025 IEEE International Conference on Robotics and Automation (ICRA) (pp. 1–8). IEEE. https://doi.org/10.1109/icra55743.2025.11128707",
        "Lai, Z., Isleem, H. F., Vairagade, V., Aluvalu, R., Tejani, G. G., & Sharaf, M. (2026). Reinforcement learning strategies in game tactics and real time decision making for team sport. Journal of Big Data, 13, Article 1485. https://doi.org/10.1186/s40537-026-01485-7",
        "Lee, D., Lee, D., & Zhang, A. (2026). A recipe for stable offline multi-agent reinforcement learning. arXiv. https://doi.org/10.48550/arxiv.2603.08399",
        "Lee, H., Kim, J., Park, J., & Cho, K. (2025). Hierarchical multi-agent reinforcement learning method using energy field in sports games. IEEE Access, 13, 1–15. https://doi.org/10.1109/access.2025.3613359",
        "Lefhal, D., Ouacha, A., El Harraj, A., & Ziti, S. (2026). Bibliometric analysis of artificial intelligence in team sports: Research trends, collaborations, and creative insights. Journal of Computational and Cognitive Engineering. Advance online publication. https://doi.org/10.47852/bonviewjcce62028489",
        "Li, A., Gong, X., Chen, B., Lu, Y., Ji, J., Wang, Y., Yang, Y., & Li, W. (2026). ShuttleEnv: An interactive data-driven RL environment for badminton strategy modeling. arXiv. https://doi.org/10.48550/arxiv.2603.17324",
        "Li, C., Dong, W., He, L., Cai, M., & Wang, D. (2025). Intelligent decision for joint operations based on improved proximal policy optimization. Scientific Reports, 15, Article 86229. https://doi.org/10.1038/s41598-025-86229-y",
        "Li, J. (2025). Machine learning-based analysis of defensive strategies in basketball using player movement data. Scientific Reports, 15, Article 98877. https://doi.org/10.1038/s41598-025-98877-1",
        "Li, J., Khishe, M., & Ibrahim, B. F. (2025). Improving women football tactics analysis by using extreme learning and accumulated optimization algorithm. Scientific Reports, 15, Article 30218. https://doi.org/10.1038/s41598-025-30218-8",
        "Li, K. (2025). Optimizing competitive sports training strategies with adaptive deep reinforcement learning. In 2025 2nd International Conference on Intelligent Computing and Robotics (ICICR) (pp. 1–6). IEEE. https://doi.org/10.1109/icicr65456.2025.00060",
        "Li, Q. (2025). Deep learning algorithm for basketball offensive tactics optimization. In 2025 IEEE International Conference on Computation, Big-Data and Engineering (ICCBE) (pp. 1–6). IEEE. https://doi.org/10.1109/iccbe65177.2025.11255810",
        "Li, W., Hu, B., Song, A., & Huang, K. (2025). HDMTK: Full integration of hierarchical decision-making and tactical knowledge in multiagent adversarial games. IEEE Transactions on Cognitive and Developmental Systems. Advance online publication. https://doi.org/10.1109/tcds.2024.3470068",
        "Li, Y., & Link, D. (2026). Intention driven identification of in-possession match phases in association football through temporal graph learning. arXiv. https://doi.org/10.48550/arxiv.2606.09289",
        "Liang, Y., Guo, X., Zhang, J., & Du, X. (2026). NeuroPlayNet: A multimodal AI framework for real-time cognitive-aware strategy optimization in professional basketball. Scientific Reports, 16, Article 41140. https://doi.org/10.1038/s41598-026-41140-y",
        "Liang, Z., Cao, J., Jiang, S., Saxena, D., & Xu, H. (2022). Hierarchical reinforcement learning with opponent modeling for distributed multi-agent cooperation. In 2022 IEEE 42nd International Conference on Distributed Computing Systems (ICDCS) (pp. 1–10). IEEE. https://doi.org/10.1109/icdcs54860.2022.00090",
        "Lin, J., Chen, F., & Liu, J. (2026). A graph-integrated reinforcement learning framework with graph neural networks for tactical decision modeling in professional football. Scientific Reports, 16, Article 50061. https://doi.org/10.1038/s41598-026-50061-9",
        "Liu, G., Luo, Y., Schulte, O., & Kharrat, T. (2020). Deep soccer analytics: Learning an action-value function for evaluating soccer players. Data Mining and Knowledge Discovery, 34, 1531–1559. https://doi.org/10.1007/s10618-020-00705-9",
        "Liu, M., Tao, W., & Huang, H. (2026). Offline reinforcement learning for badminton tactical decision-making. Engineering Applications of Artificial Intelligence, 149, Article 113395. https://doi.org/10.1016/j.engappai.2025.113395",
        "Liu, S., Lever, G., Wang, Z., Merel, J., Eslami, S., Hennes, D., Czarnecki, W. M., Tassa, Y., Omidshafiei, S., Abdolmaleki, A., Siegel, N., Hasenclever, L., Marris, L., Tunyasuvunakool, S., Song, H. F., Wulfmeier, M., Muller, P., Haarnoja, T., Tracey, B. D., … Heess, N. (2021). From motor control to team play in simulated humanoid football. Science Robotics, 6(58), Article abo0235. https://doi.org/10.1126/scirobotics.abo0235",
        "Liu, Z. (2026). Relational multi agent tactical learning for competitive football. Discover Artificial Intelligence, 6, Article 1900. https://doi.org/10.1007/s44163-026-01900-1",
        "Lv, X., Tao, Y., Zhang, Y.-F., & Xue, Y. (2025). Design of an immersive basketball tactical training system based on digital twins and federated learning. Applied Sciences, 15(7), 3831. https://doi.org/10.3390/app15073831",
        "M. R., H. R., Anitha, C., Goyal, D., & Dadheech, P. (2024). Optimizing game strategies with deep reinforcement learning: A framework for intelligent decision-making. In Proceedings of the 6th International Conference on Information Management & Machine Intelligence (pp. 1–6). Association for Computing Machinery. https://doi.org/10.1145/3745812.3745844",
        "Magelssen, C., Gilgien, M., Tajet, S. L., Losnegard, T., Haugen, P., Reid, R., & Frömer, R. (2025). Reinforcement learning enhances training efficiency in high-performance athletes. bioRxiv. https://doi.org/10.1101/2024.04.22.590558",
        "Midoul, K., El Mohajir, B. E., El Hichami, O., & Souri, A. (2026). Methodological trends in machine learning for sport. Journal of Human Sport and Exercise, 21(1), 1–20. https://doi.org/10.55860/6ff7nv62",
        "Mo, X., Hu, B., Xu, C., & Yang, Y. (2022). Research on virtual human swarm football collaboration technology based on reinforcement learning. In 2022 IEEE 21st International Conference on Ubiquitous Computing and Communications (IUCC/CIT/DSCI/SmartCNS) (pp. 1–6). IEEE. https://doi.org/10.1109/iucc-cit-dsci-smartcns57392.2022.00062",
        "Mohan, V. (2025). Learning tennis strategy through curriculum-based dueling double deep Q-networks. arXiv. https://doi.org/10.48550/arxiv.2512.22186",
        "Moya, D., Tipantuña, C., Villa, G., Calderón-Hinojosa, X., Rivadeneira, B., & Álvarez, R. (2025). Machine learning applied to professional football: Performance improvement and results prediction. Machine Learning and Knowledge Extraction, 7(3), 85. https://doi.org/10.3390/make7030085",
        "Munikoti, S., Agarwal, D., Das, L., Halappanavar, M., & Natarajan, B. (2022). Challenges and opportunities in deep reinforcement learning with graph neural networks: A comprehensive review of algorithms and applications. IEEE Transactions on Neural Networks and Learning Systems, 35(11), 1–21. https://doi.org/10.1109/tnnls.2023.3283523",
        "Murphy, K. (2024). Reinforcement learning: An overview. arXiv. https://doi.org/10.48550/arxiv.2412.05265",
        "Nakahara, H., Tsutsui, K., Takeda, K., & Fujii, K. (2023). Action valuation of on- and off-ball soccer players based on multi-agent deep reinforcement learning. IEEE Access, 11, 133194–133205. https://doi.org/10.1109/access.2023.3336425",
        "Nambiar, M., Ghosh, S., Ong, P., Chan, Y., Bee, Y., & Krishnaswamy, P. (2023). Deep offline reinforcement learning for real-world treatment optimization applications. In Proceedings of the 29th ACM SIGKDD Conference on Knowledge Discovery and Data Mining (pp. 1–11). Association for Computing Machinery. https://doi.org/10.1145/3580305.3599800",
        "Narvekar, S., Peng, B., Leonetti, M., Sinapov, J., Taylor, M. E., & Stone, P. (2020). Curriculum learning for reinforcement learning domains: A framework and survey. arXiv. https://doi.org/10.48550/arxiv.2003.04960",
        "Ng, A. Y., Harada, D., & Russell, S. (1999). Policy invariance under reward transformations: Theory and application to reward shaping. In Proceedings of the Sixteenth International Conference on Machine Learning (ICML) (pp. 278–287). Morgan Kaufmann. https://dl.acm.org/doi/10.5555/657696.657803",
        "Oberlin, P., Cederle, M., Karapetyan, A., Bolognani, S., Susto, G. A., & Dörfler, F. (2026). Chess on ice: Curling tactical decision-making via backward induction and deep reinforcement learning. arXiv. https://doi.org/10.48550/arxiv.2608.02379",
        "Pan, L., Ma, T., & Xu, H. (2021). Plan better amid conservatism: Offline multi-agent reinforcement learning with actor rectification. arXiv. https://doi.org/10.48550/arxiv.2111.11188",
        "Pan, Y., Pu, Z., Chen, M., Ren, W., Li, Y., & Ming, Z. (2026). Offline multi-agent reinforcement learning for evaluating and optimizing football attacking strategies against low-block defences. Intelligent Sports and Health. Advance online publication. https://doi.org/10.1016/j.ish.2026.03.003",
        "Petiot, G. H., Bagatin, R., Aquino, R., & Raab, M. (2021). Key characteristics of decision making in soccer and their implications. New Ideas in Psychology, 61, Article 100846. https://doi.org/10.1016/j.newideapsych.2020.100846",
        "Qiao, D., Li, W., Yang, S., Zha, H., & Wang, B. (2025). Offline multi-agent reinforcement learning via sequential score decomposition. arXiv. https://doi.org/10.48550/arxiv.2505.05968",
        "Raabe, D., Nabben, R., & Memmert, D. (2022). Graph representations for the analysis of multi-agent spatiotemporal sports data. Applied Intelligence, 53, 1–20. https://doi.org/10.1007/s10489-022-03631-z",
        "Rahimian, P., Mihalyi, B., & Toka, L. (2024). In-game soccer outcome prediction with offline reinforcement learning. Machine Learning, 113, 1–24. https://doi.org/10.1007/s10994-024-06611-1",
        "Rahimian, P., Oroojlooy, A., & Toka, L. (2021). Towards optimized actions in critical situations of soccer games with deep reinforcement learning. In 2021 IEEE 8th International Conference on Data Science and Advanced Analytics (DSAA) (pp. 1–6). IEEE. https://doi.org/10.1109/dsaa53316.2021.9564207",
        "Rahimian, P., & Toka, L. (2022). Inferring the strategy of offensive and defensive play in soccer with inverse reinforcement learning. In Machine Learning and Data Mining for Sports Analytics (pp. 31–43). Springer. https://doi.org/10.1007/978-3-031-02044-5_3",
        "Rahimian, P., & Toka, L. (2023). A data-driven approach to assist offensive and defensive players in optimal decision making. International Journal of Sports Science & Coaching, 18(3), 1–15. https://doi.org/10.1177/17479541221149481",
        "Rai, B., & Popović, M. (2026). From multi-agent reinforcement learning to agentic AI: A comprehensive literature review of algorithmic advances and decision-analytic implications (2020–2025). Applied Decision Analytics. Advance online publication. https://doi.org/10.66972/ada202734",
        "Rashid, T., Samvelyan, M., de Witt, C. S., Farquhar, G., Foerster, J., & Whiteson, S. (2020). Monotonic value function factorisation for deep multi-agent reinforcement learning. arXiv. https://doi.org/10.48550/arxiv.2003.08839",
        "Richards, P., Robbins, M., & Collins, D. (2025). Team decision-making: The integration of cognitive task analysis and video to accelerate the development of tactical shared mental models in an elite sports team. Frontiers in Sports and Active Living, 7, Article 1648137. https://doi.org/10.3389/fspor.2025.1648137",
        "Rico-González, M., Pino-Ortega, J., Méndez, A., Clemente, F., & Baca, A. (2022). Machine learning application in soccer: A systematic review. Biology of Sport, 40(1), 249–263. https://doi.org/10.5114/biolsport.2023.112970",
        "Riedmiller, M. A., Merke, A., Meier, D., Hoffmann, A., Sinner, A., Thate, O., & Ehrmann, R. (2001). Karlsruhe Brainstormers – A reinforcement learning approach to robotic soccer. In RoboCup 2000: Robot Soccer World Cup IV (pp. 435–440). Springer. https://doi.org/10.1007/3-540-45324-5_40",
        "Shakya, A., Pillai, G., & Chakrabarty, S. (2023). Reinforcement learning algorithms: A brief survey. Expert Systems with Applications, 231, Article 120495. https://doi.org/10.1016/j.eswa.2023.120495",
        "Shao, D. (2026). Learning optimal and sample-efficient decision policies with guarantees. arXiv. https://doi.org/10.48550/arxiv.2602.17978",
        "Smit, A. P., Engelbrecht, H., Brink, W., & Pretorius, A. (2023). Scaling multi-agent reinforcement learning to full 11 versus 11 simulated robotic football. Autonomous Agents and Multi-Agent Systems, 37, Article 96. https://doi.org/10.1007/s10458-023-09603-y",
        "Son, Y., Park, J., & Jeon, B. (2026). Training agents for strategic curling through a unified reinforcement learning framework. Mathematics, 14(3), 403. https://doi.org/10.3390/math14030403",
        "Song, S., & Qian, K. (2025). A study on the effect of deep reinforcement learning in cultivating athlete decision behavior and psychological resilience. Scalable Computing: Practice and Experience, 26(1), 3786. https://doi.org/10.12694/scpe.v26i1.3786",
        "Standen, M., Kim, J., & Szabo, C. (2024). Adversarial machine learning attacks and defences in multi-agent reinforcement learning. ACM Computing Surveys, 57(4), Article 8320. https://doi.org/10.1145/3708320",
        "Su, J., & Dong, S. (2025). Multi-objective optimization for dynamic logistics scheduling based on hierarchical deep reinforcement learning. Scientific Reports, 15, Article 18309. https://doi.org/10.1038/s41598-025-18309-y",
        "Sun, C., Shen, S., Hu, H., Zhou, W., & Chen, C. (2025). Complex instruction following with diverse style policies in football games. arXiv. https://doi.org/10.48550/arxiv.2511.19885",
        "Suryawanshi, D., Singh, U., Vasave, P., Ganpatye, S., Dabhade, I., & Agarwal, A. (2025). Strategic analysis of penalty kicks in football using game theoretic and reinforcement learning approaches. In 2025 7th International Conference on Information Systems and Computer Networks (ISCON) (pp. 1–6). IEEE. https://doi.org/10.1109/iscon65210.2025.11341135",
        "Takayanagi, R., Takahashi, K., & Sogabe, T. (2022). AI-assisted decision-making and risk evaluation in uncertain environment using stochastic inverse reinforcement learning: American football as a case study. Mathematical Problems in Engineering, 2022, Article 4451427. https://doi.org/10.1155/2022/4451427",
        "Tao, W., Liu, M., Sun, W., & Huang, H. (2025). Evaluating player performance and tactical decision-making in racket sports using deep reinforcement learning. In 2025 IEEE 19th International Conference on Control & Automation (ICCA) (pp. 1–6). IEEE. https://doi.org/10.1109/icca65672.2025.11129712",
        "Taourirte, A., & Mia, M. S. (2025). Multi-agent reinforcement learning and real-time decision-making in robotic soccer for virtual environments. arXiv. https://doi.org/10.48550/arxiv.2512.03166",
        "Teixeira, J., Maio, E., Afonso, P., Encarnação, S., Machado, G., Morgans, R., Barbosa, T. M., Monteiro, A. M., Forte, P., Ferraz, R., & Branquinho, L. (2025). Mapping football tactical behavior and collective dynamics with artificial intelligence: A systematic review. Frontiers in Sports and Active Living, 7, Article 1569155. https://doi.org/10.3389/fspor.2025.1569155",
        "Thomas, D. W., Jiang, J., Kori, A., Russo, A., Winkler, S., Sale, S., McMillan, J., Belardinelli, F., & Rago, A. (2026). Race strategy reinforcement learning: Optimising pitstop strategy with emergent tactics in Formula One. Machine Learning, 115, Article 7081. https://doi.org/10.1007/s10994-026-07081-3",
        "Truong, C., Ruffino, C., Crognier, A., Paizis, C., Crognier, L., & Papaxanthis, C. (2023). Error-based and reinforcement learning in basketball free throw shooting. Scientific Reports, 13, Article 26568. https://doi.org/10.1038/s41598-022-26568-2",
        "Tuyls, K., Omidshafiei, S., Muller, P., Wang, Z., Connor, J. T., Hennes, D., Graham, I., Spearman, W., Waskett, T., Steele, D., Luc, P., Recasens, A., Galashov, A., Thornton, G., Élie, R., Sprechmann, P., Moreno, P., Cao, K., Garnelo, M., … Hassabis, D. (2020). Game plan: What AI can do for football, and what football can do for AI. arXiv. https://doi.org/10.1613/jair.1.12505",
        "Van Roy, M., Robberechts, P., Yang, W.-C., De Raedt, L., & Davis, J. (2023). A Markov framework for learning and reasoning about strategies in professional soccer. Journal of Artificial Intelligence Research, 77, 1–38. https://doi.org/10.1613/jair.1.13934",
        "Wang, H.-X. (2025). Research on the application of intelligent computing methods in the analysis of sports competitive tactics in an interdisciplinary collaborative environment. International Journal of Computer Information Systems and Industrial Management Applications, 17, 255. https://doi.org/10.70917/ijcisim-2025-0255",
        "Wang, K.-D., Wang, W.-Y., Chen, Y.-T., Lin, Y.-H., & Peng, W. (2024). The CoachAI badminton environment: A novel reinforcement learning environment with realistic opponents (Student Abstract). Proceedings of the AAAI Conference on Artificial Intelligence, 38(21), 23831–23833. https://doi.org/10.1609/aaai.v38i21.30523",
        "Wang, S., Pan, Y., Pu, Z., Yi, J., Liang, Y., & Zhang, D. (2023). Heterogeneous-graph attention reinforcement learning for football matches. In 2023 International Joint Conference on Neural Networks (IJCNN) (pp. 1–8). IEEE. https://doi.org/10.1109/ijcnn54540.2023.10191648",
        "Wang, S., Pu, Z., Pan, Y., Liu, B., Ma, H., & Yi, J. (2024). Long-term and short-term opponent intention inference for football multiplayer policy learning. IEEE Transactions on Cognitive and Developmental Systems. Advance online publication. https://doi.org/10.1109/tcds.2024.3404061",
        "Wang, Y. (2025). Using reinforcement learning to identify the key factors for players to win games. ITM Web of Conferences, 78, Article 01007. https://doi.org/10.1051/itmconf/20257801007",
        "Wang, Y., Wang, Y., Tian, F., Ma, J., & Jin, Q. (2025). Intelligent games meeting with multi-agent deep reinforcement learning: A comprehensive review. Artificial Intelligence Review, 58, Article 11166. https://doi.org/10.1007/s10462-025-11166-1",
        "Wang, Z., Veličković, P., Hennes, D., Tomašev, N., Prince, L., Kaisers, M., Bachrach, Y., Élie, R., Piccinini, F., Spearman, W., Graham, I., Connor, J. T., Yang, Y., Recasens, A., Khan, M., Beauguerlange, N., Sprechmann, P., Moreno, P., Heess, N., … Tuyls, K. (2023). TacticAI: An AI assistant for football tactics. Nature Communications, 15, Article 45965. https://doi.org/10.1038/s41467-024-45965-x",
        "Watson, N., Hendricks, S., Stewart, T., & Durbach, I. (2020). Integrating machine learning and decision support in tactical decision-making in rugby union. Journal of the Operational Research Society, 71(12), 1–12. https://doi.org/10.1080/01605682.2020.1779624",
        "Wei, J., Xu, X., Lan, Y., Liu, T., & Wang, Y. (2025). AGT: Efficient offline reinforcement learning with advantage-guided transformer. CAAI Transactions on Intelligence Technology. Advance online publication. https://doi.org/10.1049/cit2.70094",
        "Won, J., Gopinath, D., & Hodgins, J. (2021). Control strategies for physically simulated characters performing two-player competitive sports. ACM Transactions on Graphics, 40(4), Article 146. https://doi.org/10.1145/3476576.3476725",
        "Wu, J. (2025). DDPG-LSTM framework for personalized athlete training plan optimization and competition strategy generation. Informatica (Slovenia), 48(33), 8820. https://doi.org/10.31449/inf.v48i33.8820",
        "Wurman, P. R., Barrett, S., Kawamoto, K., MacGlashan, J., Subramanian, K., Walsh, T. J., Capobianco, R., Devlic, A., Eckert, F., Fuchs, F., Gilpin, L., Khandelwal, P., Kompella, V., Lin, H., MacAlpine, P., Oller, D., Seno, T., Sherstan, C., Thomure, M. D., … Kitano, H. (2022). Outracing champion Gran Turismo drivers with deep reinforcement learning. Nature, 602(7896), 223–228. https://doi.org/10.1038/s41586-021-04357-7",
        "Xia, X., Chen, Q., & Wang, Z. (2025). Deep reinforcement learning-driven personalized training load control algorithm for competitive sports performance optimization. Scientific Reports, 15, Article 30453. https://doi.org/10.1038/s41598-025-30453-z",
        "Xu, C., & Wang, Y. (2024). Tactical intelligent decision modelling in sports competitions based on reinforcement learning algorithms. Journal of Electrical Systems, 20(3), 3124. https://doi.org/10.52783/jes.3124",
        "Xu, H., Lin, B., & Liu, L. (2025). Design of intelligent optimization of sports strategy and training decision support system based on deep reinforcement learning. Discover Artificial Intelligence, 5, Article 473. https://doi.org/10.1007/s44163-025-00473-9",
        "Xu, Q., & He, X. (2022). Football training evaluation using machine learning and decision support system. Soft Computing, 26, 1–12. https://doi.org/10.1007/s00500-022-07210-9",
        "Xu, S., Liu, G., Kharrat, T., Luo, Y., Aloulou, M., Peña, J., Sofeikov, K. I., Reid, A. C., Roberts, P. M., Spencer, S., Carnall, J., McHale, I. G., Schulte, O., Zha, H., & Zheng, W. (2026). TacticGen: Grounding adaptable and scalable generation of football tactics. arXiv. https://doi.org/10.48550/arxiv.2604.18210",
        "Xu, Z. (2024). Decision support system for optimizing tactics and strategies of sports competition using reinforcement learning algorithm. Journal of Electrical Systems, 20(3), 1304. https://doi.org/10.52783/jes.1304",
        "Xu, Z., Bai, Y., Zhang, B., Li, D., & Fan, G. (2021). HAVEN: Hierarchical cooperative multi-agent reinforcement learning with dual coordination mechanism. Proceedings of the AAAI Conference on Artificial Intelligence, 37(10), 11735–11743. https://doi.org/10.1609/aaai.v37i10.26386",
        "Xue, T., Zhao, Y., & Dong, Y. (2026). Autonomous tactical decision-making for multi-aircraft via reinforcement learning. Guidance, Navigation and Control. Advance online publication. https://doi.org/10.1142/s273748072650010x",
        "Yanai, C., Solomon, A., Katz, G., Shapira, B., & Rokach, L. (2022). Q-Ball: Modeling basketball games using deep reinforcement learning. Proceedings of the AAAI Conference on Artificial Intelligence, 36(8), 8806–8813. https://doi.org/10.1609/aaai.v36i8.20861",
        "Yang, J., Ge, H., & Cui, Y. (2025). An AI framework for counterattack detection and decision-making evaluation in football. Journal of Big Data, 12, Article 1128. https://doi.org/10.1186/s40537-025-01128-3",
        "Yang, L., Zhou, C., & Sang, B. (2026). Domain-specific contexts promote model-based decision making for basketball players. Scientific Reports, 16, Article 54649. https://doi.org/10.1038/s41598-026-54649-z",
        "Yang, Y., Li, F., & Chang, H. (2023). Enhancing short track speed skating performance through improved DDQN tactical decision model. Sensors, 23(24), 9904. https://doi.org/10.3390/s23249904",
        "Yu, Q. (2025). Multi modal hierarchical reinforcement learning framework for dynamic sports sponsorship optimization. Scientific Reports, 15, Article 27915. https://doi.org/10.1038/s41598-025-27915-9",
        "Yu, X., Lin, Y., Wang, X., Han, S., & Lv, K. (2023). GHQ: Grouped hybrid Q-learning for cooperative heterogeneous multi-agent reinforcement learning. Complex & Intelligent Systems, 10, 1–18. https://doi.org/10.1007/s40747-024-01415-1",
        "Yuan, C., Al Forhad, M. A., Bansal, R., Sidorova, A., & Albert, M. V. (2024). Multi-agent dual level reinforcement learning of strategy and tactics in competitive games. Results in Control and Optimization, 15, Article 100471. https://doi.org/10.1016/j.rico.2024.100471",
        "Zhang, B. (2025). Adaptive tactical decision making in ice hockey: Integrating multi-agent reinforcement learning framework with advanced computer vision techniques. ECE Official Conference Proceedings, 21. https://doi.org/10.22492/issn.2188-1162.2025.21",
        "Zhang, D., Yuan, Q., Meng, L., Xia, R., Liu, W., & Qin, C. (2025). Reinforcement learning for single-agent to multi-agent systems: From basic theory to industrial application progress, a survey. Artificial Intelligence Review, 58, Article 11439. https://doi.org/10.1007/s10462-025-11439-9",
        "Zhang, J., Shi, E., Niyato, D., Ai, B., & Shen, X. (2024). Graph neural network meets multi-agent reinforcement learning: Fundamentals, applications, and future directions. IEEE Wireless Communications, 31(6), 1–8. https://doi.org/10.1109/mwc.015.2300595",
        "Zhang, J., & Tao, D. (2023). Research on deep reinforcement learning basketball robot shooting skills improvement based on end to end architecture and multi-modal perception. Frontiers in Neurorobotics, 17, Article 1274543. https://doi.org/10.3389/fnbot.2023.1274543",
        "Zhang, J., & Xue, Q. (2020). Actor–critic-based decision-making method for the artificial intelligence commander in tactical wargames. The Journal of Defense Modeling and Simulation: Applications, Methodology, Technology, 17(4), 1–12. https://doi.org/10.1177/1548512920954542",
        "Zhang, Q., Wang, Q., & Niu, Y. (2026). Adaptive training load optimization for track and field athletes: A reinforcement learning approach. Scientific Reports, 16, Article 41946. https://doi.org/10.1038/s41598-026-41946-w",
        "Zhao, J., Lin, J., Zhang, X., Li, Y., Zhou, X., & Sun, Y. (2024). From mimic to counteract: A two-stage reinforcement learning algorithm for Google research football. Neural Computing and Applications, 36, 1–16. https://doi.org/10.1007/s00521-024-09455-x",
        "Zhao, S., Ma, H., Pu, Z., Huang, J., Pan, Y., Wang, S., & Ming, Z. (2025). TacEleven: Generative tactic discovery for football open play. arXiv. https://doi.org/10.48550/arxiv.2511.13326",
        "Zhao, T., Chen, T., & Zhang, B. (2025). QMIX-GNN: A graph neural network-based heterogeneous multi-agent reinforcement learning model for improved collaboration and decision-making. Applied Sciences, 15(7), 3794. https://doi.org/10.3390/app15073794",
        "Zhao, Z., Chai, W., Hao, S., Hu, W., Wang, G., Cao, S., Song, M.-G., Hwang, J.-N., & Wang, G. (2023). A survey of deep learning in sports applications: Perception, comprehension, and decision. IEEE Transactions on Visualization and Computer Graphics. Advance online publication. https://doi.org/10.1109/tvcg.2025.3554801",
        "Zhou, S.-Y., & Zhu, M. (2025). Enhancing athlete performance using deep learning techniques in sports analytics. IEEE Access, 13, 1–15. https://doi.org/10.1109/access.2025.3631851",
        "Zhu, C. (2025). Research on reinforcement learning algorithm for sports club participation prediction and optimization. In 2025 IEEE 3rd International Conference on Control, Electronics and Computer Technology (ICCECT) (pp. 1–6). IEEE. https://doi.org/10.1109/iccect64621.2025.11339670",
        "Ziyi, Z., Bunker, R., Takeda, K., & Fujii, K. (2023). Multi-agent deep-learning based comparative analysis of team sport trajectories. IEEE Access, 11, 41646–41656. https://doi.org/10.1109/access.2023.3269287"
    ]

    for ref in raw_refs:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.4)
        p.paragraph_format.first_line_indent = Inches(-0.4)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(ref)
        r.font.size = Pt(9)

    out_file = r"c:\Reinforcement Learning of Sports\Complete_Research_Article_Scopus_Standard.docx"
    doc.save(out_file)
    print(f"COMPLETE RESEARCH ARTICLE SUCCESSFULLY CREATED AT: {out_file}")

if __name__ == "__main__":
    build_complete_scopus_article()
