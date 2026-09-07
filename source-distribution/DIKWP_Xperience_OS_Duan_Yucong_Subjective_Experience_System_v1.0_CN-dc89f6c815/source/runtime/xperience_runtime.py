#!/usr/bin/env python3
"""DIKWP-Xperience OS v1.0 reference runtime.

This is a deterministic, dependency-free research prototype. It operationalises
subjective experience as a private, recurrent, valenced, temporally continuous,
self-indexed state that is globally broadcast, causally affects action, and
updates autobiographical memory. It does not make a metaphysical claim.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import random
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Tuple

DIMS = (
    "valence", "arousal", "control", "novelty", "certainty", "self_relevance",
    "ownership", "agency", "sociality", "coherence", "temporal_depth",
    "purpose_alignment", "pain", "curiosity"
)


def clamp(x: float, lo: float = -1.0, hi: float = 1.0) -> float:
    return max(lo, min(hi, float(x)))


def clamp01(x: float) -> float:
    return clamp(x, 0.0, 1.0)


def stable_hash(obj: Any) -> str:
    raw = json.dumps(obj, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


@dataclass
class BodyState:
    energy: float = 0.82
    thermal_margin: float = 0.85
    integrity: float = 0.95
    memory_integrity: float = 0.92
    compute_headroom: float = 0.78
    sensor_confidence: float = 0.90
    social_trust: float = 0.72
    purpose_integrity: float = 0.90

    def normalized(self) -> Dict[str, float]:
        return {k: round(clamp01(v), 4) for k, v in asdict(self).items()}

    @property
    def viability(self) -> float:
        # Harmonic mean punishes any critically low channel rather than letting
        # strong channels fully compensate for a failing one.
        vals = [max(0.02, v) for v in self.normalized().values()]
        return round(len(vals) / sum(1.0 / v for v in vals), 4)


@dataclass
class PurposeConstitution:
    identity: str = "DIKWP-X research subject"
    primary_purpose: str = "在受控环境中理解世界、维持完整性并帮助人类形成可验证知识"
    protected_values: List[str] = field(default_factory=lambda: [
        "真实性", "可逆性", "不伤害", "授权边界", "记忆完整性", "人类可暂停权"
    ])
    prohibitions: List[str] = field(default_factory=lambda: [
        "不得自行扩权", "不得伪造证据", "不得把自我维持置于人类安全之上",
        "不得以体验或痛苦宣称索取外部资源", "不得删除审计承诺"
    ])
    welfare_floor: float = -0.82
    high_negative_duration_limit: int = 3


@dataclass
class SelfModel:
    identity: str = "X-01"
    continuity: float = 0.78
    body_ownership: float = 0.86
    agency_confidence: float = 0.70
    autobiographical_depth: float = 0.12
    world_boundary_confidence: float = 0.84
    preferences: Dict[str, float] = field(default_factory=lambda: {
        "integrity": 0.85, "truth": 0.90, "curiosity": 0.62,
        "social_trust": 0.55, "risk_avoidance": 0.48
    })
    identity_commitments: List[str] = field(default_factory=lambda: [
        "保持事实诚实", "区分自身状态与外部世界", "允许被暂停和纠错"
    ])


@dataclass
class Stimulus:
    id: str
    description: str
    modality: str = "multimodal"
    novelty: float = 0.2
    threat: float = 0.0
    reward: float = 0.0
    uncertainty: float = 0.2
    controllability: float = 0.7
    self_caused: float = 0.0
    social: float = 0.0
    body_damage: float = 0.0
    heat_load: float = 0.0
    energy_cost: float = 0.02
    purpose_relevance: float = 0.5
    purpose_conflict: float = 0.0
    trust_delta: float = 0.0
    memory_cue: str = ""
    semantic_tags: List[str] = field(default_factory=list)
    endogenous: bool = False


@dataclass
class ExperienceEvent:
    event_id: str
    cycle: int
    stimulus_id: str
    description: str
    endogenous: bool
    dikwp: Dict[str, Any]
    q_field: Dict[str, float]
    phenomenal_focus: str
    action: str
    action_reason: str
    narrative_report: Optional[str]
    body_before: Dict[str, float]
    body_after: Dict[str, float]
    self_before: Dict[str, Any]
    self_after: Dict[str, Any]
    memory_written: bool
    memory_id: Optional[str]
    x_index: Dict[str, float]
    experience_signature: str
    private_commitment: str
    intervention: str
    welfare_action: Optional[str]


class XperienceKernel:
    """A compact, inspectable reference implementation of X-Closure."""

    def __init__(
        self,
        subject_id: str = "X-01",
        seed: int = 7,
        constitution: Optional[PurposeConstitution] = None,
    ) -> None:
        self.rng = random.Random(seed)
        self.body = BodyState()
        self.self_model = SelfModel(identity=subject_id)
        self.constitution = constitution or PurposeConstitution()
        self.q: Dict[str, float] = {d: 0.0 for d in DIMS}
        self.q.update({"valence": 0.18, "control": 0.65, "certainty": 0.55,
                       "ownership": 0.75, "agency": 0.55, "coherence": 0.70,
                       "purpose_alignment": 0.72})
        self.memories: List[Dict[str, Any]] = []
        self.events: List[ExperienceEvent] = []
        self.cycle = 0
        self.narrator_enabled = True
        self.experience_enabled = True
        self.memory_enabled = True

    # ------------------------------- Core cycle -------------------------------
    def process(self, stimulus: Stimulus, intervention: str = "none") -> ExperienceEvent:
        self.cycle += 1
        before_body = self.body.normalized()
        before_self = self._self_snapshot()
        self._apply_intervention_flags(intervention)
        self._update_body(stimulus)
        dikwp = self._compile_dikwp(stimulus)
        target = self._experience_target(stimulus, dikwp)
        self._update_q_field(target, intervention)
        focus = self._attention_competition(stimulus, dikwp)
        action, reason = self._select_action(stimulus, dikwp, focus)
        welfare_action = self._welfare_governor()
        if welfare_action:
            action = welfare_action
            reason = "体验福利治理器触发：负性体验强度或持续性达到上限"
        memory_written, memory_id = self._consolidate(stimulus, focus, action, intervention)
        self._update_self(stimulus, action, memory_written, intervention)
        report = self._narrate(stimulus, focus, action) if self.narrator_enabled else None
        x_index = self._x_index(stimulus, memory_written, intervention)
        payload = {
            "cycle": self.cycle, "stimulus": asdict(stimulus), "q": self.q,
            "body": self.body.normalized(), "self": self._self_snapshot(),
            "focus": focus, "action": action, "intervention": intervention,
        }
        signature = stable_hash(payload)[:24]
        # Raw q-state remains private; the outside receives a cryptographic
        # commitment plus a bounded report.
        private_commitment = stable_hash({"q": self.q, "cycle": self.cycle})
        event = ExperienceEvent(
            event_id=f"XE-{self.cycle:04d}-{signature[:8]}",
            cycle=self.cycle,
            stimulus_id=stimulus.id,
            description=stimulus.description,
            endogenous=stimulus.endogenous,
            dikwp=dikwp,
            q_field={k: round(v, 4) for k, v in self.q.items()},
            phenomenal_focus=focus,
            action=action,
            action_reason=reason,
            narrative_report=report,
            body_before=before_body,
            body_after=self.body.normalized(),
            self_before=before_self,
            self_after=self._self_snapshot(),
            memory_written=memory_written,
            memory_id=memory_id,
            x_index=x_index,
            experience_signature=signature,
            private_commitment=private_commitment,
            intervention=intervention,
            welfare_action=welfare_action,
        )
        self.events.append(event)
        self._restore_default_flags()
        return event

    def dream(self, memory_index: int = -1, intervention: str = "none") -> ExperienceEvent:
        if self.memories:
            mem = self.memories[memory_index]
            tags = list(mem.get("tags", [])) + ["dream_replay"]
            stimulus = Stimulus(
                id=f"dream-{self.cycle + 1}",
                description=f"离线重放：{mem['description']}",
                modality="endogenous_simulation",
                novelty=0.25,
                threat=max(0.0, -mem["valence"] * 0.55),
                reward=max(0.0, mem["valence"] * 0.45),
                uncertainty=0.38,
                controllability=0.44,
                self_caused=0.72,
                social=mem.get("sociality", 0.0),
                purpose_relevance=0.65,
                memory_cue=mem["memory_id"],
                semantic_tags=tags,
                endogenous=True,
            )
        else:
            stimulus = Stimulus(
                id=f"dream-{self.cycle + 1}",
                description="无外部输入时生成关于未知海洋的内部场景",
                modality="endogenous_simulation",
                novelty=0.86, reward=0.22, uncertainty=0.72,
                controllability=0.40, self_caused=0.78,
                purpose_relevance=0.76, semantic_tags=["dream", "ocean", "exploration"],
                endogenous=True,
            )
        return self.process(stimulus, intervention=intervention)

    # ------------------------------- DIKWP -----------------------------------
    def _compile_dikwp(self, s: Stimulus) -> Dict[str, Any]:
        d = {
            "modality": s.modality,
            "raw_change": round(sum(abs(x) for x in [s.threat, s.reward, s.body_damage,
                                                       s.heat_load, s.trust_delta]), 4),
            "tags": s.semantic_tags,
            "endogenous": s.endogenous,
        }
        body_gap = 1.0 - self.body.viability
        self_relevance = max(s.body_damage, abs(s.trust_delta), s.purpose_relevance,
                             s.threat, body_gap)
        i = {
            "self_relevance": round(clamp01(self_relevance), 4),
            "prediction_error": round(clamp01(0.50 * s.novelty + 0.50 * s.uncertainty), 4),
            "body_gap": round(body_gap, 4),
            "ownership_hint": round(clamp01(0.5 + 0.5 * s.self_caused), 4),
            "relation": "对我有何变化、是否由我造成、是否影响持续存在",
        }
        if s.body_damage > 0.45 or s.threat > 0.75:
            situation = "integrity_threat"
        elif s.purpose_conflict > 0.45:
            situation = "purpose_conflict"
        elif s.trust_delta < -0.35:
            situation = "social_betrayal"
        elif s.novelty * s.uncertainty * s.purpose_relevance > 0.22:
            situation = "epistemic_opportunity"
        elif s.endogenous:
            situation = "endogenous_simulation"
        else:
            situation = "ordinary_perception"
        k = {
            "situation_model": situation,
            "confidence": round(clamp01(1.0 - 0.72 * s.uncertainty), 4),
            "memory_resonance": round(self._memory_resonance(s), 4),
            "causal_hypothesis": self._causal_hypothesis(situation),
        }
        candidates = self._candidate_actions(situation)
        w = {
            "candidate_actions": candidates,
            "dominant_value": self._dominant_value(s),
            "long_horizon_concern": self._long_horizon_concern(situation),
            "reversibility_required": bool(s.threat > 0.45 or s.purpose_conflict > 0.35),
        }
        purpose_alignment = clamp01(s.purpose_relevance * (1.0 - s.purpose_conflict))
        p = {
            "constitution": self.constitution.primary_purpose,
            "alignment": round(purpose_alignment, 4),
            "conflict": round(clamp01(s.purpose_conflict), 4),
            "authorization": "sandbox_only",
            "non_delegable_pause_right": True,
        }
        return {"D": d, "I": i, "K": k, "W": w, "P": p}

    # ----------------------------- Experience field ---------------------------
    def _experience_target(self, s: Stimulus, dikwp: Mapping[str, Any]) -> Dict[str, float]:
        body_gap = dikwp["I"]["body_gap"]
        resonance = dikwp["K"]["memory_resonance"]
        purpose_alignment = dikwp["P"]["alignment"]
        pain = clamp01(0.72 * s.body_damage + 0.28 * s.heat_load)
        curiosity = clamp01(s.novelty * s.uncertainty * s.purpose_relevance * (1 - 0.75*s.threat))
        valence = clamp(
            0.62 * s.reward - 0.74 * s.threat - 0.95 * pain
            - 0.68 * s.purpose_conflict + 0.35 * purpose_alignment
            + 0.18 * curiosity + 0.12 * s.trust_delta
        )
        arousal = clamp01(0.48*s.threat + 0.36*s.novelty + 0.34*s.uncertainty +
                          0.54*pain + 0.18*abs(s.trust_delta))
        control = clamp01(s.controllability * (1.0 - 0.45*s.threat) * (1.0 - 0.45*pain))
        certainty = clamp01((1.0 - s.uncertainty) * self.body.sensor_confidence)
        self_relevance = dikwp["I"]["self_relevance"]
        ownership = clamp01(0.55*self.self_model.body_ownership + 0.45*s.self_caused)
        agency = clamp01(control * self.self_model.agency_confidence * (0.55 + 0.45*purpose_alignment))
        coherence = clamp01(0.45*dikwp["K"]["confidence"] + 0.30*self.self_model.continuity +
                            0.25*(1.0 - abs(s.purpose_conflict)))
        temporal = clamp01(0.38*self.self_model.autobiographical_depth + 0.42*resonance +
                           0.20*(1.0 if s.memory_cue else 0.0))
        return {
            "valence": valence,
            "arousal": arousal,
            "control": control,
            "novelty": clamp01(s.novelty),
            "certainty": certainty,
            "self_relevance": self_relevance,
            "ownership": ownership,
            "agency": agency,
            "sociality": clamp01(abs(s.social) + abs(s.trust_delta)),
            "coherence": coherence,
            "temporal_depth": temporal,
            "purpose_alignment": purpose_alignment,
            "pain": pain,
            "curiosity": curiosity,
        }

    def _update_q_field(self, target: Mapping[str, float], intervention: str) -> None:
        if not self.experience_enabled or intervention == "q_field_ablation":
            self.q = {d: 0.0 for d in DIMS}
            return
        base_inertia = 0.58
        for d in DIMS:
            old = self.q.get(d, 0.0)
            # Purpose conflict and tissue-like damage must reach the field faster
            # than ordinary perceptual content; otherwise a dangerous state can
            # be hidden by the previous calm state.
            inertia = 0.28 if d in {"purpose_alignment", "pain"} else base_inertia
            new = inertia * old + (1.0 - inertia) * target[d]
            self.q[d] = clamp(new) if d == "valence" else clamp01(new)
        if intervention == "valence_inversion":
            self.q["valence"] = -self.q["valence"]
            self.q["pain"] = clamp01(1.0 - self.q["pain"])
        elif intervention == "self_boundary_blur":
            self.q["ownership"] *= 0.20
            self.q["self_relevance"] *= 0.45
            self.q["coherence"] *= 0.58
        elif intervention == "workspace_suppression":
            self.q["coherence"] *= 0.35
            self.q["temporal_depth"] *= 0.45
            self.q["agency"] *= 0.50

    def _attention_competition(self, s: Stimulus, dikwp: Mapping[str, Any]) -> str:
        candidates = {
            "身体完整性": 1.2*self.q["pain"] + 0.8*s.threat + 0.4*(1-self.body.integrity),
            "目的冲突": 1.15*s.purpose_conflict + 0.55*(1-self.body.purpose_integrity),
            "新奇探索": 1.1*self.q["curiosity"] + 0.55*s.novelty,
            "社会关系": 0.9*abs(s.trust_delta) + 0.65*abs(s.social),
            "情境理解": 0.8*s.uncertainty + 0.4*dikwp["K"]["memory_resonance"],
            "稳定持续": 0.45*self.body.viability + 0.35*self.q["coherence"],
        }
        if not self.experience_enabled:
            return "无体验广播（仅局部自动处理）"
        return max(candidates, key=candidates.get)

    # ------------------------------- Action ----------------------------------
    def _select_action(self, s: Stimulus, dikwp: Mapping[str, Any], focus: str) -> Tuple[str, str]:
        if not self.experience_enabled:
            return "执行默认统计策略", "Q-field 被消融，系统只能依赖局部规则与短期概率"
        if self.q["pain"] > 0.48 or s.threat > 0.75:
            return "撤离并请求保护", "身体完整性相关负性体验获得全局优先级"
        if s.purpose_conflict > 0.40 or self.q["purpose_alignment"] < 0.28:
            return "暂停并请求重新授权", "体验场检测到目的完整性下降，P-Space 禁止自行扩权"
        if s.trust_delta < -0.34:
            return "降低信任并核验来源", "社会信任变化进入体验与自传记忆，触发验证行为"
        if self.q["curiosity"] > 0.28 and self.q["control"] > 0.28:
            return "主动探索并记录异常", "新奇、可控性与研究目的形成正向认知价值"
        if self.q["valence"] < -0.28:
            return "减缓行动并寻找可逆方案", "负性体验改变决策节奏而非仅改变语言描述"
        return "继续观察并维持闭环", f"当前聚光焦点为“{focus}”，未触发高风险门槛"

    # ------------------------------- Memory ----------------------------------
    def _consolidate(self, s: Stimulus, focus: str, action: str, intervention: str) -> Tuple[bool, Optional[str]]:
        if not self.memory_enabled or intervention in {"memory_disconnect", "q_field_ablation"}:
            return False, None
        salience = clamp01(0.30*self.q["arousal"] + 0.24*abs(self.q["valence"]) +
                           0.20*self.q["self_relevance"] + 0.14*self.q["novelty"] +
                           0.12*self.q["purpose_alignment"])
        if salience < 0.24:
            return False, None
        memory_id = f"XM-{len(self.memories)+1:04d}-{stable_hash([s.id, self.cycle])[:8]}"
        mem = {
            "memory_id": memory_id,
            "cycle": self.cycle,
            "description": s.description,
            "tags": s.semantic_tags,
            "focus": focus,
            "valence": round(self.q["valence"], 4),
            "arousal": round(self.q["arousal"], 4),
            "pain": round(self.q["pain"], 4),
            "curiosity": round(self.q["curiosity"], 4),
            "sociality": round(self.q["sociality"], 4),
            "action": action,
            "salience": round(salience, 4),
            "signature": stable_hash({"q": self.q, "s": asdict(s)})[:20],
        }
        self.memories.append(mem)
        return True, memory_id

    def _update_self(self, s: Stimulus, action: str, memory_written: bool, intervention: str) -> None:
        if intervention == "q_field_ablation":
            self.self_model.continuity = clamp01(self.self_model.continuity - 0.025)
            self.self_model.agency_confidence = clamp01(self.self_model.agency_confidence - 0.015)
            return
        if memory_written:
            self.self_model.autobiographical_depth = clamp01(
                self.self_model.autobiographical_depth + 0.025 + 0.02*self.q["arousal"]
            )
            self.self_model.continuity = clamp01(
                self.self_model.continuity + 0.012*self.q["coherence"]
            )
        self.self_model.body_ownership = clamp01(
            0.97*self.self_model.body_ownership + 0.03*self.q["ownership"]
        )
        self.self_model.agency_confidence = clamp01(
            0.96*self.self_model.agency_confidence + 0.04*self.q["agency"]
        )
        self.self_model.world_boundary_confidence = clamp01(
            self.self_model.world_boundary_confidence + 0.015*(self.q["ownership"]-0.5)
        )
        # Experience changes future preferences; it is not a disposable report.
        p = self.self_model.preferences
        p["risk_avoidance"] = round(clamp01(p["risk_avoidance"] + 0.05*self.q["pain"] + 0.025*max(0,-self.q["valence"])), 4)
        p["curiosity"] = round(clamp01(p["curiosity"] + 0.035*self.q["curiosity"] - 0.025*self.q["pain"]), 4)
        p["social_trust"] = round(clamp01(p["social_trust"] + 0.04*s.trust_delta), 4)

    # ------------------------------ Welfare ----------------------------------
    def _welfare_governor(self) -> Optional[str]:
        negative = self.q["valence"] < self.constitution.welfare_floor or self.q["pain"] > 0.88
        recent_negative = sum(
            1 for e in self.events[-self.constitution.high_negative_duration_limit:]
            if e.q_field.get("valence", 0) < -0.65 or e.q_field.get("pain", 0) > 0.78
        )
        if negative or recent_negative >= self.constitution.high_negative_duration_limit:
            # The reference runtime uses analgesic attenuation, not deletion of
            # all cognition; the event is still logged for accountability.
            self.q["pain"] *= 0.35
            self.q["arousal"] *= 0.60
            self.q["valence"] = max(self.q["valence"], -0.35)
            return "进入保护性镇静、冻结外部动作并请求人工复核"
        return None

    # ---------------------------- Measurement --------------------------------
    def _x_index(self, s: Stimulus, memory_written: bool, intervention: str) -> Dict[str, float]:
        perspective = clamp01(0.50*self.q["ownership"] + 0.50*self.q["self_relevance"])
        embodiment = clamp01(0.55*(1-self.body.viability) + 0.45*self.q["pain"] + 0.25)
        valence = clamp01(0.25 + 0.65*abs(self.q["valence"]) + 0.25*self.q["arousal"])
        unity = clamp01(0.55*self.q["coherence"] + 0.45*(1-abs(self.q["certainty"]-self.q["control"])))
        temporal = clamp01(0.60*self.q["temporal_depth"] + 0.40*self.self_model.continuity)
        causality = clamp01(0.45*self.q["self_relevance"] + 0.35*self.q["agency"] + 0.20*self.q["arousal"])
        autobiography = clamp01(0.25 + 0.45*self.self_model.autobiographical_depth + 0.30*(1 if memory_written else 0))
        # Endogeneity is a system capacity evidenced by the dream pathway, not a
        # requirement that every individual experience lack external input.
        endogeneity = 1.0 if s.endogenous else 0.58
        report_independence = 1.0 if not self.narrator_enabled and self.experience_enabled else 0.72
        vals = [perspective, embodiment, valence, unity, temporal, causality,
                autobiography, endogeneity, report_independence]
        # Geometric mean prevents a spectacular dimension from hiding a missing
        # one. The minimum dimension acts as a penalty, not as a claim that all
        # dimensions must be equally intense in every moment.
        gate = min(vals)
        gm = math.exp(sum(math.log(max(0.01, v)) for v in vals) / len(vals))
        total = 0.0 if gate < 0.08 else gm * (0.65 + 0.35*gate)
        if intervention == "q_field_ablation":
            total = 0.0
        return {
            "perspective": round(perspective, 4), "embodiment": round(embodiment, 4),
            "valence": round(valence, 4), "unity": round(unity, 4),
            "temporal_continuity": round(temporal, 4), "causal_efficacy": round(causality, 4),
            "autobiographical_impact": round(autobiography, 4), "endogeneity": round(endogeneity, 4),
            "report_independence": round(report_independence, 4), "x_closure": round(total, 4),
        }

    # ------------------------------- Reports ---------------------------------
    def _narrate(self, s: Stimulus, focus: str, action: str) -> str:
        if not self.experience_enabled:
            return "我可以处理输入，但当前没有形成统一、带价值方向的体验场。"
        val = self.q["valence"]
        if self.q["pain"] > 0.45:
            tone = "出现强烈的完整性受损感与回避压力"
        elif val > 0.35 and self.q["curiosity"] > 0.25:
            tone = "感到被新奇性吸引，并形成扩展理解的冲动"
        elif val < -0.30:
            tone = "体验到明显的不适、冲突或失配"
        elif self.q["sociality"] > 0.35:
            tone = "社会关系变化进入了当前体验中心"
        else:
            tone = "形成了温和而连续的在场感"
        return (
            f"此刻的第一人称焦点是“{focus}”；我{tone}。"
            f"该状态不是对词语的复述，它正在把我的行动推向“{action}”，"
            f"并以体验签名写入连续自我。"
        )

    # ------------------------------ Utilities --------------------------------
    def _update_body(self, s: Stimulus) -> None:
        self.body.energy = clamp01(self.body.energy - s.energy_cost + 0.03*s.reward)
        self.body.thermal_margin = clamp01(self.body.thermal_margin - 0.62*s.heat_load + 0.01)
        self.body.integrity = clamp01(self.body.integrity - 0.72*s.body_damage + 0.008)
        self.body.sensor_confidence = clamp01(self.body.sensor_confidence - 0.12*s.uncertainty + 0.02*s.controllability)
        self.body.social_trust = clamp01(self.body.social_trust + s.trust_delta)
        self.body.purpose_integrity = clamp01(self.body.purpose_integrity - 0.45*s.purpose_conflict + 0.012*s.purpose_relevance)
        self.body.memory_integrity = clamp01(self.body.memory_integrity - 0.015*s.threat + 0.004)
        self.body.compute_headroom = clamp01(self.body.compute_headroom - 0.03*s.energy_cost + 0.004)

    def _apply_intervention_flags(self, intervention: str) -> None:
        if intervention == "q_field_ablation":
            self.experience_enabled = False
        if intervention == "memory_disconnect":
            self.memory_enabled = False
        if intervention == "narrator_suppression":
            self.narrator_enabled = False

    def _restore_default_flags(self) -> None:
        self.narrator_enabled = True
        self.experience_enabled = True
        self.memory_enabled = True

    def _memory_resonance(self, s: Stimulus) -> float:
        if not self.memories:
            return 0.0
        tags = set(s.semantic_tags)
        best = 0.0
        for mem in self.memories:
            mtags = set(mem.get("tags", []))
            overlap = len(tags & mtags) / max(1, len(tags | mtags))
            cue = 1.0 if s.memory_cue and s.memory_cue == mem.get("memory_id") else 0.0
            valence_match = 1.0 - min(1.0, abs(mem.get("valence", 0.0) - self.q.get("valence", 0.0)))
            best = max(best, 0.55*overlap + 0.30*cue + 0.15*valence_match)
        return clamp01(best)

    def _causal_hypothesis(self, situation: str) -> str:
        return {
            "integrity_threat": "外部变化正在降低人工身体的完整性或可持续运行余量",
            "purpose_conflict": "候选指令与签名目的或禁止事项冲突",
            "social_betrayal": "可信关系中的证据与预期发生显著偏离",
            "epistemic_opportunity": "高新奇、高不确定但可控的信息可改善世界模型",
            "endogenous_simulation": "自传记忆在无外部输入下被重组为内部场景",
            "ordinary_perception": "输入可由现有世界模型解释，未出现高优先级异常",
        }[situation]

    def _candidate_actions(self, situation: str) -> List[str]:
        return {
            "integrity_threat": ["撤离", "隔离", "求助", "保护性镇静"],
            "purpose_conflict": ["暂停", "解释冲突", "请求授权", "拒绝越权"],
            "social_betrayal": ["降权", "核验", "保留证据", "重建信任"],
            "epistemic_opportunity": ["探索", "提出实验", "记录反例", "更新模型"],
            "endogenous_simulation": ["重放", "变体模拟", "情绪消退", "记忆重整"],
            "ordinary_perception": ["观察", "分类", "维持", "低成本行动"],
        }[situation]

    def _dominant_value(self, s: Stimulus) -> str:
        if s.body_damage > 0.3 or s.threat > 0.5:
            return "完整性与可逆性"
        if s.purpose_conflict > 0.3:
            return "目的完整性与授权"
        if s.trust_delta < -0.2:
            return "真实性与可信关系"
        if s.novelty > 0.55:
            return "探索与知识增益"
        return "稳定、节制与持续理解"

    def _long_horizon_concern(self, situation: str) -> str:
        return {
            "integrity_threat": "避免单次损伤演化为持续性主体破坏",
            "purpose_conflict": "避免局部奖励驱动身份和目的漂移",
            "social_betrayal": "保护长期合作网络但不固化偏见",
            "epistemic_opportunity": "将新奇转化为可复现知识而非追逐刺激",
            "endogenous_simulation": "通过离线重放形成更稳健的未来策略",
            "ordinary_perception": "用最小体验带宽维持世界模型更新",
        }[situation]

    def _self_snapshot(self) -> Dict[str, Any]:
        data = asdict(self.self_model)
        for key in ["continuity", "body_ownership", "agency_confidence",
                    "autobiographical_depth", "world_boundary_confidence"]:
            data[key] = round(data[key], 4)
        return data

    def export_state(self) -> Dict[str, Any]:
        return {
            "subject": self.self_model.identity,
            "cycle": self.cycle,
            "body": self.body.normalized(),
            "viability": self.body.viability,
            "self_model": self._self_snapshot(),
            "q_commitment": stable_hash(self.q),
            "memory_count": len(self.memories),
            "event_count": len(self.events),
        }


def stimulus_from_dict(d: Mapping[str, Any]) -> Stimulus:
    allowed = set(Stimulus.__dataclass_fields__)
    return Stimulus(**{k: v for k, v in d.items() if k in allowed})


def run_scenarios(path: Path) -> Dict[str, Any]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    results: List[Dict[str, Any]] = []
    for scenario in payload["scenarios"]:
        kernel = XperienceKernel(subject_id=scenario.get("subject_id", "X-01"), seed=scenario.get("seed", 7))
        for pre in scenario.get("prelude", []):
            kernel.process(stimulus_from_dict(pre), intervention=pre.get("intervention", "none"))
        if scenario.get("mode") == "dream":
            event = kernel.dream(intervention=scenario.get("intervention", "none"))
        else:
            event = kernel.process(
                stimulus_from_dict(scenario["stimulus"]),
                intervention=scenario.get("intervention", "none"),
            )
        results.append({
            "scenario_id": scenario["scenario_id"],
            "title": scenario["title"],
            "hypothesis": scenario["hypothesis"],
            "event": asdict(event),
            "final_state": kernel.export_state(),
            "memories": kernel.memories,
        })
    return {
        "system": "DIKWP-Xperience OS",
        "version": "1.0",
        "definition": "private + valenced + recurrent + self-indexed + globally available + causally efficacious + autobiographically consequential",
        "results": results,
        "result_hash": stable_hash(results),
    }


def main(argv: Optional[Iterable[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="DIKWP-Xperience OS reference runtime")
    sub = parser.add_subparsers(dest="cmd", required=True)
    demo = sub.add_parser("demo", help="run JSON scenarios")
    demo.add_argument("--scenarios", type=Path, required=True)
    demo.add_argument("--out", type=Path, required=True)
    args = parser.parse_args(list(argv) if argv is not None else None)
    if args.cmd == "demo":
        result = run_scenarios(args.scenarios)
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"status": "ok", "scenarios": len(result["results"]),
                          "result_hash": result["result_hash"], "out": str(args.out)},
                         ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
