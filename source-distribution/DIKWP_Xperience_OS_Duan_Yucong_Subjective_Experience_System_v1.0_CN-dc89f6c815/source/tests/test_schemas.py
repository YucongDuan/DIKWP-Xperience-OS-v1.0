import importlib.util
import json
import sys
import unittest
from pathlib import Path

import jsonschema

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("xruntime", ROOT / "runtime" / "xperience_runtime.py")
mod = importlib.util.module_from_spec(spec)
sys.modules["xruntime"] = mod
spec.loader.exec_module(mod)


class SchemaTests(unittest.TestCase):
    def test_body_schema(self):
        schema = json.loads((ROOT / "schemas" / "synthetic_body_state.schema.json").read_text())
        jsonschema.validate(mod.BodyState().normalized(), schema)

    def test_constitution_schema(self):
        schema = json.loads((ROOT / "schemas" / "purpose_constitution.schema.json").read_text())
        instance = json.loads((ROOT / "config" / "purpose_constitution.example.json").read_text())
        jsonschema.validate(instance, schema)

    def test_event_schema(self):
        schema = json.loads((ROOT / "schemas" / "experience_event.schema.json").read_text())
        s = mod.Stimulus(id="demo", description="demo", novelty=.8, reward=.2,
                         uncertainty=.6, purpose_relevance=.8, semantic_tags=["demo"])
        e = mod.XperienceKernel().process(s)
        jsonschema.validate(mod.asdict(e), schema)

    def test_genome_schema(self):
        schema = json.loads((ROOT / "schemas" / "experience_genome.schema.json").read_text())
        instance = json.loads((ROOT / "config" / "default_experience_genome.json").read_text())
        jsonschema.validate(instance, schema)

    def test_intervention_schema(self):
        schema = json.loads((ROOT / "schemas" / "experience_intervention.schema.json").read_text())
        instance = {
            "intervention_id": "XI-TEST-001",
            "type": "narrator_suppression",
            "authorization": "dual-review-test-key",
            "target": "X-01/FirstPersonNarrator",
            "expected_effect": "报告被关闭，Q-field、行动与记忆保持",
            "recovery_plan": "恢复叙述器并比较前后承诺哈希",
            "timestamp": "2026-07-13T00:00:00Z",
            "result_commitment": None
        }
        jsonschema.validate(instance, schema)


if __name__ == "__main__":
    unittest.main()
