# System Architecture & Workflow Diagrams
## Multi-Agent Reinforcement Learning for Football Simulation

This document provides visual architectural, geometric, and experimental workflow diagrams representing the research protocol and system implementation.

---

### Diagram 1: End-to-End System Architecture (CTDE MAPPO Pipeline)
```mermaid
flowchart TB
    subgraph SIM[" Google Research Football (GRF) Simulator "]
        GRF_ENV["Physics Engine & Match State (p, v, ball)"]
    end

    subgraph PERCEPTION[" Perception & Feature Engineering "]
        RAW_OBS["Raw Cartesian Observations\n[x, y, vx, vy]"]
        
        subgraph SEMANTIC_ENGINE[" Vectorized Semantic Feature Engine (Algorithm 1) "]
            EDMS["Space Score S(j)\nk1=-0.5, k2=0.3, k3=0.4"]
            L_PASS["Pass-Lane Occlusion L_pass\nGaussian Ray Corridor"]
            TTR["Time-To-Reach (TTR)\n& Pitch Control"]
            SHOT["Shot Viability Score\n(Surrogate xG)"]
        end

        AUG_OBS["Augmented Agent Observation\no_aug = [o_raw, F_sem]"]
    end

    subgraph DECENTRALIZED_ACTORS[" Decentralized Execution (Actors) "]
        ACTOR_1["Agent 1 Policy π_θ\n(Ball Carrier)"]
        ACTOR_2["Agent 2 Policy π_θ\n(Off-Ball Runner)"]
        ACTOR_N["Agent N Policy π_θ\n(Off-Ball Runner)"]
    end

    subgraph REWARD_SYSTEM[" Tactical Valuation & Reward Engine (Algorithm 2) "]
        SPARSE_R["Sparse Goal Reward\n(+1.0 / -1.0)"]
        XT_GRID["Markovian xT Grid (16x12)\nΔOBV Progression"]
        SPACE_R["Off-Ball Space Creation\nmax(0, ΔS(j))"]
        DISRUPT_R["Defensive Disruption\nDefender Drag Vector Δp_d"]
        TOTAL_R["Composite Reward\nR_total = w_sp*r_sp + w_obv*ΔOBV + w_sp*r_sp + w_dis*r_dis"]
    end

    subgraph CENTRALIZED_CRITIC[" Centralized Training (Critic & Optimization) "]
        GLOBAL_S["True Global Match State (s_t)"]
        CRITIC["Centralized Critic V_ϕ(s_t)\nState-Value Estimation"]
        GAE["Generalized Advantage Estimation\nGAE(γ=0.993, λ=0.95)"]
        PPO_OPT["MAPPO Clipped Optimization\nPolicy Loss + Value Loss + Entropy"]
    end

    %% Data Connections
    GRF_ENV --> RAW_OBS
    GRF_ENV --> GLOBAL_S
    RAW_OBS --> SEMANTIC_ENGINE
    EDMS --> AUG_OBS
    L_PASS --> AUG_OBS
    TTR --> AUG_OBS
    SHOT --> AUG_OBS
    RAW_OBS --> AUG_OBS

    AUG_OBS --> ACTOR_1
    AUG_OBS --> ACTOR_2
    AUG_OBS --> ACTOR_N

    ACTOR_1 -->|Discrete Actions a_t| GRF_ENV
    ACTOR_2 -->|Discrete Actions a_t| GRF_ENV
    ACTOR_N -->|Discrete Actions a_t| GRF_ENV

    GRF_ENV -->|Transition s_{t+1}| REWARD_SYSTEM
    SPARSE_R --> TOTAL_R
    XT_GRID --> TOTAL_R
    SPACE_R --> TOTAL_R
    DISRUPT_R --> TOTAL_R

    GLOBAL_S --> CRITIC
    TOTAL_R --> GAE
    CRITIC --> GAE
    GAE --> PPO_OPT
    PPO_OPT -.->|Policy Gradient Update| ACTOR_1
    PPO_OPT -.->|Policy Gradient Update| ACTOR_2
    PPO_OPT -.->|Policy Gradient Update| ACTOR_N
    PPO_OPT -.->|Value Gradient Update| CRITIC
```

---

### Diagram 2: Geometric Tactical Feature Calculation (Pass-Lane & Space Score)
```mermaid
flowchart LR
    subgraph GEOMETRY[" Vectorized Geometry on Normalized Pitch [-1, 1] x [-0.42, 0.42] "]
        direction TB
        B["Ball Carrier b (p_b)"]
        J["Teammate Receiver j (p_j)"]
        D["Defender d (p_d)"]
        
        V_PASS["Pass Vector: v_pass = p_j - p_b"]
        PROJ["Projection: t_proj = (p_d - p_b) · v_pass / ||v_pass||²"]
        CLOSEST["Closest Trajectory Point:\np_closest = p_b + clip(t_proj, 0, 1) * v_pass"]
        ORTHO["Orthogonal Distance: h_perp = ||p_d - p_closest||"]
        GAUSS["Gaussian Dynamic Reach:\nP_intercept = exp(-h_perp² / (2*σ_d²))"]
        L_OPEN["Pass Openness:\nL_pass = 1.0 - max(P_intercept)"]
        
        B --> V_PASS
        J --> V_PASS
        V_PASS --> PROJ
        D --> PROJ
        PROJ --> CLOSEST
        CLOSEST --> ORTHO
        D --> ORTHO
        ORTHO --> GAUSS
        GAUSS --> L_OPEN
    end

    subgraph SYNTHESIS[" Tactical Space Score Synthesis "]
        D_BALL["Normalized Ball Distance:\nD_ball = ||p_j - p_b|| / D_max"]
        D_DEF["Nearest Defender Distance:\nD_def = min ||p_j - p_d|| / D_safe"]
        
        SPACE_FORMULA["Space Score S(j):\nS(j) = -0.50 * D_ball + 0.30 * D_def + 0.40 * L_pass"]
        
        D_BALL --> SPACE_FORMULA
        D_DEF --> SPACE_FORMULA
        L_OPEN --> SPACE_FORMULA
    end
```

---

### Diagram 3: 4-Phase Research & Comparative Experimental Protocol
```mermaid
flowchart TD
    subgraph PHASE1[" Phase 1: Environment & Data Calibration "]
        GRF_SETUP["Google Research Football\n(Academy 3v1 / Pass&Shoot / 11v11)"]
        STATSBOMB["StatsBomb Open Data (140+ Leagues)\nPitch Transition Calibration"]
        XT_MATRIX["Precomputed 16x12 xT Matrix\n(Expected Threat / OBV Surface)"]
        GRF_SETUP --- STATSBOMB --> XT_MATRIX
    end

    subgraph PHASE2[" Phase 2: 2x2 Factorial Ablation Matrix "]
        M1["M1 (Control Baseline):\nRaw Obs + Sparse Reward"]
        M2["M2 (State Only):\nSemantic Obs + Sparse Reward"]
        M3["M3 (Reward Only):\nRaw Obs + Tactical Composite Reward"]
        M4["M4 (Full Novel System):\nSemantic Obs + Tactical Composite Reward"]
    end

    subgraph PHASE3[" Phase 3: Multi-Seed Controlled Training "]
        SEEDS["5 Independent Random Seeds per Model\n(Seeds: 42, 101, 2024, 7, 888)"]
        TRAIN_EXEC["Parallel PPO Rollouts (K=16 Envs)\nTotal Step Budget: 3M - 5M steps/seed"]
        WNB["Telemetry Logging (W&B / TensorBoard)\nPolicy Entropy, GAE Value Loss, Returns"]
        SEEDS --> TRAIN_EXEC --> WNB
    end

    subgraph PHASE4[" Phase 4: Tactical Benchmarking & Interpretability "]
        TPCA["Tactical Phase Classification Accuracy\n(Target: >89% TPCA)"]
        OBMQ["Off-Ball Movement Quality\n(Normalized Space Index: ~0.90)"]
        COHEN["Decision Coherence (Cohen's κ > 0.70)\nCoaching Rule Accordance"]
        RISK_MOD["Game Context Risk Modulation\nTrailing (36% Through-balls) vs Leading (18%)"]
        STATS["Welch's t-test (p < 0.01)\nEffect Size (Cohen's d > 1.2) + 95% CI"]
    end

    PHASE1 --> PHASE2
    PHASE2 --> PHASE3
    PHASE3 --> PHASE4
```

---

### Diagram 4: Pitch Coordinate System and Tactical Feature Geometry

```
Normalized Pitch Dimensions: X in [-1.0, 1.0], Y in [-0.42, 0.42]

+-----------------------------------+-----------------------------------+
| Defensive Half                    | Attacking Half                    |
|                                   |                                   |
|                                   |               [Receiver j]        |
|                                   |              / (p_j)              |
|                                   |             /                     |
|                                   |   Gaussian /                      |
|                                   |   Corridor/                       |
|                                   |         |/                        |
|                                   |      [Defender d]                 |
|                                   |      (h_perp distance)            |
|                                   |        /                          |
|                                   |       / v_pass                    |
|        [Own Goal]                 |      /                  [Opp Goal]|
|        [-1.0, 0.0]        (0.0, 0.0)   /                    [+1.0,0.0]|
|            |                      |   [Ball Carrier b]          |     |
|            |                      |   (p_b)                     |     |
|                                   |                                   |
|                                   |                                   |
|                                   |                                   |
+-----------------------------------+-----------------------------------+

Key Geometric Variables:
  * v_pass = p_j - p_b (Passing vector connecting carrier to receiver)
  * t_proj = Scalar projection along v_pass
  * h_perp = Orthogonal Euclidean distance from defender d to passing line
  * P_intercept = exp( -h_perp^2 / (2 * sigma_d^2) )
  * L_pass = 1.0 - max_{d} P_intercept
  * Space Score S(j) = -0.50 * D_ball + 0.30 * D_def + 0.40 * L_pass
```
