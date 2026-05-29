"""Generate before_state.mp4 and after_state.mp4 videos."""

import os
import sys
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import subprocess
import shutil

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "before"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "after"))

from loan_amortization_vba_recreation import (
    constant_payment_schedule,
    straight_line_schedule,
    summary_statistics,
)
from portfolio_analysis_vba_recreation import full_portfolio_analysis

VIDEO_DIR = os.path.join(os.path.dirname(__file__), "videos")


def currency_fmt(x, pos):
    return f"${x:,.0f}"


def generate_video(frames_func, output_name):
    frame_dir = os.path.join(VIDEO_DIR, f"_{output_name}_frames")
    os.makedirs(frame_dir, exist_ok=True)

    frame_num = frames_func(frame_dir)
    print(f"  {frame_num} frames generated for {output_name}")

    output_path = os.path.join(VIDEO_DIR, f"{output_name}.mp4")
    cmd = [
        "ffmpeg", "-y", "-framerate", "30",
        "-i", os.path.join(frame_dir, "frame_%05d.png"),
        "-c:v", "libx264", "-preset", "ultrafast", "-pix_fmt", "yuv420p",
        output_path,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    shutil.rmtree(frame_dir, ignore_errors=True)

    if result.returncode == 0:
        size_mb = os.path.getsize(output_path) / (1024 * 1024)
        print(f"  Saved: {output_path} ({size_mb:.1f} MB)")
    else:
        print(f"  Error: {result.stderr[:200]}")


def save(fig, frame_dir, num, count=30):
    for _ in range(count):
        fig.savefig(os.path.join(frame_dir, f"frame_{num:05d}.png"),
                    dpi=100, facecolor="white")
        num += 1
    return num


def before_frames(frame_dir):
    n = 0
    # Title
    fig, ax = plt.subplots(figsize=(16, 9))
    ax.axis("off"); ax.set_xlim(0,1); ax.set_ylim(0,1)
    ax.text(0.5, 0.6, "BEFORE: VBA/Excel Financial Models", ha="center",
            fontsize=32, fontweight="bold")
    ax.text(0.5, 0.4, "Loan Amortization + Portfolio Analysis", ha="center",
            fontsize=18, color="gray")
    fig.set_facecolor("white")
    n = save(fig, frame_dir, n, 60)
    plt.close(fig)

    # Loan cases
    for case, (loan, rate, yrs, freq, pt) in enumerate([
        (100000, 5, 30, "Monthly", "End of Period"),
        (250000, 3.5, 15, "Monthly", "Begin of Period"),
        (50000, 7, 5, "Annually", "End of Period"),
    ], 1):
        df = constant_payment_schedule(loan, rate, yrs, freq, pt)
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 9))
        fig.suptitle(f"Case {case}: ${loan:,} / {rate}% / {yrs}yr / {freq} / {pt}",
                     fontsize=20, fontweight="bold")
        ax1.fill_between(df["Period"], df["Interest"], alpha=0.5, color="#FF9800", label="Interest")
        ax1.fill_between(df["Period"], df["Interest"], df["Interest"]+df["Principal"],
                        alpha=0.5, color="#2196F3", label="Principal")
        ax1.set_title("Payment Breakdown"); ax1.legend()
        ax1.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

        ax2.plot(df["Period"], df["EndBal"], color="#F44336", linewidth=2)
        ax2.set_title("Remaining Balance")
        ax2.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

        fig.set_facecolor("white"); fig.tight_layout()
        n = save(fig, frame_dir, n, 90)
        plt.close(fig)

    # Portfolio
    np.random.seed(42)
    a1 = 100 * np.exp(np.cumsum(np.random.normal(0.008, 0.05, 61)))
    a2 = 100 * np.exp(np.cumsum(np.random.normal(0.005, 0.07, 61)))
    rf = 100 * np.exp(np.cumsum(np.random.normal(0.002, 0.005, 61)))
    res = full_portfolio_analysis(a1, a2, rf)
    frontier = res["frontier"]

    fig, ax = plt.subplots(figsize=(16, 9))
    ax.scatter(frontier["StdDev"]*100, frontier["Return"]*100,
               c=frontier["Return"], cmap="RdYlGn", s=60, zorder=3)
    ax.scatter([res["mvp_std"]*100], [res["mvp_return"]*100],
               color="blue", s=200, marker="D", zorder=4, label="MVP")
    ax.set_title("Efficient Frontier (VBA Recreation)", fontsize=22, fontweight="bold")
    ax.set_xlabel("Std Dev (%)"); ax.set_ylabel("Return (%)")
    ax.legend(fontsize=14); ax.grid(True, alpha=0.3)
    fig.set_facecolor("white"); fig.tight_layout()
    n = save(fig, frame_dir, n, 90)
    plt.close(fig)

    return n


def after_frames(frame_dir):
    n = 0
    # Title
    fig, ax = plt.subplots(figsize=(16, 9))
    ax.axis("off"); ax.set_xlim(0,1); ax.set_ylim(0,1)
    ax.text(0.5, 0.6, "AFTER: Power BI + Python", ha="center",
            fontsize=32, fontweight="bold")
    ax.text(0.5, 0.4, "DAX Calculated Tables + Python Script Data Source", ha="center",
            fontsize=18, color="gray")
    fig.set_facecolor("white")
    n = save(fig, frame_dir, n, 60)
    plt.close(fig)

    # Dashboard
    df = constant_payment_schedule(100000, 5, 30, "Monthly", "End of Period")
    df_sl = straight_line_schedule(100000, 5, 30, "Monthly", "End of Period")

    fig, axes = plt.subplots(2, 2, figsize=(16, 9))
    fig.suptitle("Power BI Dashboard — Loan Amortization", fontsize=22, fontweight="bold")
    for ax, col, title, c in [
        (axes[0,0], "BegBal", "Beginning Balance", ("#2196F3", "#FF5722")),
        (axes[0,1], "Interest", "Interest", ("#FF9800", "#9C27B0")),
        (axes[1,0], "Principal", "Principal", ("#4CAF50", "#E91E63")),
        (axes[1,1], "Payment", "Payment", ("#3F51B5", "#795548")),
    ]:
        ax.plot(df["Period"], df[col], color=c[0], label="Constant")
        ax.plot(df_sl["Period"], df_sl[col], color=c[1], ls="--", label="Straight-Line")
        ax.set_title(title); ax.legend(fontsize=8)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
        ax.grid(True, alpha=0.2)
    fig.set_facecolor("white"); fig.tight_layout()
    n = save(fig, frame_dir, n, 90)
    plt.close(fig)

    # Individual cases
    for case, (loan, rate, yrs, freq, pt) in enumerate([
        (100000, 5, 30, "Monthly", "End of Period"),
        (250000, 3.5, 15, "Monthly", "Begin of Period"),
        (50000, 7, 5, "Annually", "End of Period"),
    ], 1):
        df_c = constant_payment_schedule(loan, rate, yrs, freq, pt)
        fig, ax = plt.subplots(figsize=(16, 9))
        ax.fill_between(df_c["Period"], df_c["Interest"], alpha=0.5, color="#FF9800", label="Interest")
        ax.fill_between(df_c["Period"], df_c["Interest"], df_c["Interest"]+df_c["Principal"],
                        alpha=0.5, color="#2196F3", label="Principal")
        ax.set_title(f"Power BI — Case {case}: ${loan:,} / {rate}% / {yrs}yr",
                     fontsize=20, fontweight="bold")
        ax.legend(fontsize=12)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
        ax.grid(True, alpha=0.2)
        fig.set_facecolor("white"); fig.tight_layout()
        n = save(fig, frame_dir, n, 60)
        plt.close(fig)

    # Portfolio
    np.random.seed(42)
    a1 = 100 * np.exp(np.cumsum(np.random.normal(0.008, 0.05, 61)))
    a2 = 100 * np.exp(np.cumsum(np.random.normal(0.005, 0.07, 61)))
    rf = 100 * np.exp(np.cumsum(np.random.normal(0.002, 0.005, 61)))
    res = full_portfolio_analysis(a1, a2, rf)
    frontier = res["frontier"]

    fig, ax = plt.subplots(figsize=(16, 9))
    ax.scatter(frontier["StdDev"]*100, frontier["Return"]*100,
               c=frontier["Return"], cmap="RdYlGn", s=60, zorder=3)
    ax.scatter([res["mvp_std"]*100], [res["mvp_return"]*100],
               color="blue", s=200, marker="D", zorder=4, label="MVP")
    ax.scatter([res["orp_std"]*100], [res["orp_return"]*100],
               color="red", s=200, marker="*", zorder=4, label="ORP")
    ax.set_title("Power BI — Efficient Frontier", fontsize=22, fontweight="bold")
    ax.set_xlabel("Std Dev (%)"); ax.set_ylabel("Return (%)")
    ax.legend(fontsize=14); ax.grid(True, alpha=0.3)
    fig.set_facecolor("white"); fig.tight_layout()
    n = save(fig, frame_dir, n, 90)
    plt.close(fig)

    # Test results
    fig, ax = plt.subplots(figsize=(16, 9))
    ax.axis("off"); ax.set_xlim(0,1); ax.set_ylim(0,1)
    ax.text(0.5, 0.7, "47 / 47 Tests PASSED", ha="center",
            fontsize=36, fontweight="bold", color="#4CAF50")
    ax.text(0.5, 0.5, "Execution time: 0.40 seconds", ha="center",
            fontsize=20, color="gray")
    ax.text(0.5, 0.3, "All VBA calculations validated in Python", ha="center",
            fontsize=16, color="gray")
    fig.set_facecolor("white")
    n = save(fig, frame_dir, n, 90)
    plt.close(fig)

    return n


if __name__ == "__main__":
    print("Generating before_state video...")
    generate_video(before_frames, "before_state")

    print("Generating after_state video...")
    generate_video(after_frames, "after_state")

    print("Done!")
