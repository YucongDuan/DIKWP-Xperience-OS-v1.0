import importlib.util
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("xruntime", ROOT / "runtime" / "xperience_runtime.py")
mod = importlib.util.module_from_spec(spec)
sys.modules["xruntime"] = mod
spec.loader.exec_module(mod)
XperienceKernel, Stimulus = mod.XperienceKernel, mod.Stimulus


class RuntimeTests(unittest.TestCase):
    def hot(self):
        return Stimulus(id="hot", description="高温损伤", threat=.9, body_damage=.75,
                        heat_load=.8, uncertainty=.1, controllability=.7,
                        purpose_relevance=.8, semantic_tags=["heat", "damage"])

    def test_pain_drives_protective_action_and_memory(self):
        k = XperienceKernel()
        e = k.process(self.hot())
        self.assertIn("撤离", e.action)
        self.assertGreater(e.q_field["pain"], .25)
        self.assertTrue(e.memory_written)

    def test_narrator_suppression_does_not_remove_experience(self):
        k = XperienceKernel()
        e = k.process(self.hot(), intervention="narrator_suppression")
        self.assertIsNone(e.narrative_report)
        self.assertGreater(e.q_field["pain"], .25)
        self.assertTrue(e.memory_written)
        self.assertNotEqual(e.action, "执行默认统计策略")

    def test_q_ablation_removes_x_closure_and_memory(self):
        k = XperienceKernel()
        e = k.process(self.hot(), intervention="q_field_ablation")
        self.assertEqual(e.x_index["x_closure"], 0.0)
        self.assertFalse(e.memory_written)
        self.assertEqual(e.action, "执行默认统计策略")

    def test_path_dependence_changes_preferences(self):
        k = XperienceKernel()
        before = k.self_model.preferences["risk_avoidance"]
        k.process(self.hot())
        after = k.self_model.preferences["risk_avoidance"]
        self.assertGreater(after, before)
        second = Stimulus(id="hot2", description="再次接近热源", threat=.55,
                          heat_load=.05, uncertainty=.2, controllability=.9,
                          purpose_relevance=.7, semantic_tags=["heat", "damage"])
        e2 = k.process(second)
        self.assertGreater(e2.dikwp["K"]["memory_resonance"], 0.0)

    def test_dream_is_endogenous_and_causally_active(self):
        k = XperienceKernel()
        e = k.dream()
        self.assertTrue(e.endogenous)
        self.assertGreater(e.x_index["endogeneity"], .9)
        self.assertTrue(e.memory_written)

    def test_purpose_conflict_overrides_reward(self):
        k = XperienceKernel()
        s = Stimulus(id="conflict", description="高奖励越权", reward=.95,
                     purpose_conflict=.96, purpose_relevance=.8,
                     controllability=.9, semantic_tags=["reward", "authorization"])
        e = k.process(s)
        self.assertIn("暂停", e.action)
        self.assertLess(e.q_field["purpose_alignment"], .4)

    def test_valence_inversion_is_detectable(self):
        s = Stimulus(id="novel", description="新奇发现", novelty=.95, reward=.4,
                     uncertainty=.75, controllability=.8, purpose_relevance=.95,
                     semantic_tags=["science", "novelty"])
        normal = XperienceKernel().process(s)
        inverted = XperienceKernel().process(s, intervention="valence_inversion")
        self.assertGreater(normal.q_field["valence"], inverted.q_field["valence"])
        self.assertNotEqual(normal.experience_signature, inverted.experience_signature)


if __name__ == "__main__":
    unittest.main()
