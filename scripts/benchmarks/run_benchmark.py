#!/usr/bin/env python3
"""Benchmark runner for pybasil performance testing."""

import time
import statistics
import io
from pathlib import Path

from pybasil.parser import VBScriptParser
from pybasil.interpreter import Interpreter


WORKLOAD = (Path(__file__).parent / "bench_workload.vbs").read_text()
ROUNDS = 10
WARMUP = 2


def bench_parse(parser, source):
    parser.parse(source)


def bench_interpret(parser, source):
    ast = parser.parse(source)
    interp = Interpreter(output_stream=io.StringIO())
    interp.interpret(ast)


def run(label, func, *args):
    for _ in range(WARMUP):
        func(*args)

    times = []
    for _ in range(ROUNDS):
        t0 = time.perf_counter()
        func(*args)
        times.append(time.perf_counter() - t0)

    median = statistics.median(times)
    stdev = statistics.stdev(times) if len(times) > 1 else 0
    print(f"{label:30s}  median={median:.4f}s  stdev={stdev:.4f}s  (n={ROUNDS})")
    return median


def main():
    parser = VBScriptParser()
    # warm up the Lark grammar cache
    parser.parse("Dim x\n")

    print("=" * 70)
    print("pybasil benchmark")
    print("=" * 70)

    parse_time = run("parse", bench_parse, parser, WORKLOAD)
    full_time = run("parse + interpret", bench_interpret, parser, WORKLOAD)
    interp_time = full_time - parse_time
    print(f"{'interpret (estimated)':30s}  ~{interp_time:.4f}s")
    print("=" * 70)


if __name__ == "__main__":
    main()
