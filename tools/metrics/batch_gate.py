# -*- coding: utf-8 -*-
"""Batched iteration gate: several fixes, one run, automatic attribution.

  python batch_gate.py "<census expr>" --exe NEW.exe --base BASE.exe
         [--flags OXI_S1589_DISABLE,OXI_S1590_DISABLE] [--jobs 3] [--identity 12]

What it does, compared with subset_gate.py:
  * runs documents in PARALLEL (--jobs worker processes; each renderer is capped
    at OXI_MEM_CAP_MB, so 3 fits a 14GB machine);
  * CACHES the base binary's result per document under
    pipeline_data/gate_cache/<sha12 of BASE>/ -- a base is measured once, ever;
  * for every PASS->FAIL document, re-runs that document once per --flags entry
    (the fix's opt-out env) and reports which flag restores the PASS, so a batch
    of fixes needs one gate run and at most (#regressions x #flags) re-renders.

Full gates remain the commit rule; this is the fast loop.
"""
import os, sys, json, random, subprocess, hashlib
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO / "tools" / "metrics"))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import feature_census as FC  # noqa: E402
from subset_gate import truth_for  # noqa: E402

CACHE = REPO / "pipeline_data" / "gate_cache"
ONE = r'''
import json, os, sys
sys.path.insert(0, sys.argv[1])
import measure_pagination_oxi as MO, pagination_diff as PD
word = json.load(open(sys.argv[3], encoding="utf-8"))
d = PD.diff_doc("x", word, MO.measure_doc(sys.argv[2]))
print(json.dumps({"pass": d["pass"], "score": d["score"], "pcd": d.get("page_count_delta")}))
'''


def sha12(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()[:12]


def run_one(exe, path, truth, extra_env=None):
    env = dict(os.environ, OXI_GDI_EXE=exe)
    if extra_env:
        env.update(extra_env)
    r = subprocess.run([sys.executable, "-c", ONE, str(REPO / "tools" / "metrics"), path, str(truth)],
                       capture_output=True, text=True, env=env)
    try:
        return json.loads(r.stdout.strip().splitlines()[-1])
    except Exception:
        return {"pass": None, "err": (r.stderr or r.stdout)[-200:]}


def dump(exe, path):
    import tempfile, shutil
    tmp = tempfile.mkdtemp(prefix="bg_")
    try:
        out = os.path.join(tmp, "l.json")
        subprocess.run([exe, path, os.path.join(tmp, "p"), "110", "--dump-layout=" + out], capture_output=True)
        return open(out, encoding="utf-8").read() if os.path.exists(out) else None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    expr = sys.argv[1]
    a = sys.argv[2:]
    arg = lambda k, d=None: a[a.index(k) + 1] if k in a else d
    exe, base = arg("--exe"), arg("--base")
    flags = [f for f in (arg("--flags", "") or "").split(",") if f]
    jobs = int(arg("--jobs", "3"))
    n_id = int(arg("--identity", "0"))
    rows = FC.load()
    hits = FC.query(expr)
    work = [(d, rows[d]["path"], truth_for(d)) for d in hits]
    work = [w for w in work if w[2] is not None]
    print(f"predicate: {expr}\nmatched {len(hits)} docs, {len(work)} with truth, jobs={jobs}", flush=True)

    bdir = CACHE / sha12(base)
    bdir.mkdir(parents=True, exist_ok=True)

    def base_result(w):
        did, path, truth = w
        f = bdir / (did.replace("/", "__") + ".json")
        if f.exists():
            return json.loads(f.read_text(encoding="utf-8"))
        r = run_one(base, path, truth)
        if r.get("pass") is not None:
            f.write_text(json.dumps(r), encoding="utf-8")
        return r

    with ThreadPoolExecutor(max_workers=jobs) as pool:
        new_res = list(pool.map(lambda w: run_one(exe, w[1], w[2]), work))
        base_res = list(pool.map(base_result, work))

    flips = []
    n_new = n_base = 0
    for w, n, b in zip(work, new_res, base_res):
        n_new += n.get("pass") is True
        n_base += b.get("pass") is True
        tag = ""
        if b.get("pass") is True and n.get("pass") is not True:
            tag = "  <<< PASS->FAIL"; flips.append(w)
        elif b.get("pass") is not True and n.get("pass") is True:
            tag = "  >>> FAIL->PASS"
        if tag or n.get("pass") is not True:
            print(f"  {w[0]}: base={b.get('pass')} {b.get('score')}  new={n.get('pass')} {n.get('score')}{tag}")
    print(f"NEW pass {n_new}/{len(work)}   BASE pass {n_base}/{len(work)}", flush=True)

    if flips and flags:
        print("attribution (flag that restores PASS):")
        tasks = [(w, fl) for w in flips for fl in flags]
        with ThreadPoolExecutor(max_workers=jobs) as pool:
            outs = list(pool.map(lambda t: run_one(exe, t[0][1], t[0][2], {t[1]: "1"}), tasks))
        for (w, fl), r in zip(tasks, outs):
            print(f"  {w[0]} with {fl}: pass={r.get('pass')} {r.get('score')}")

    if n_id:
        rest = [d for d in rows if d not in set(hits) and "error" not in rows[d]]
        random.seed(0)
        sample = random.sample(rest, min(n_id, len(rest)))
        with ThreadPoolExecutor(max_workers=jobs) as pool:
            pairs = list(pool.map(lambda d: (d, dump(exe, rows[d]["path"]) == dump(base, rows[d]["path"])), sample))
        bad = [d for d, same in pairs if not same]
        for d in bad:
            print(f"  identity {d}: DIFFERS <<<")
        print(f"identity sample: {len(sample) - len(bad)}/{len(sample)} byte-identical")


if __name__ == "__main__":
    main()
