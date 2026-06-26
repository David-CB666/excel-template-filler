#!/usr/bin/env python3
"""Hero image for excel-template-filler: dual-engine architecture conceptual diagram."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import matplotlib.patches as mpatches
from pathlib import Path

def render(out_path):
    fig, ax = plt.subplots(figsize=(12, 5.5), dpi=150)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 5.5)
    ax.axis('off')
    
    # Colors
    DARK = '#1F3864'
    BLUE = '#2E75B6'
    LIGHT_BLUE = '#D6E4F0'
    GREEN = '#2D8B4A'
    ORANGE = '#ED7D31'
    GRAY = '#808080'
    WHITE = '#FFFFFF'
    BG = '#F8F9FB'
    
    fig.patch.set_facecolor(BG)
    
    # --- Title ---
    ax.text(0.5, 5.15, 'Excel Template Filler', fontsize=22, fontweight='bold',
           color=DARK, ha='left')
    ax.text(0.55, 4.85, 'Dual-Engine Architecture', fontsize=12, color='#666',
           ha='left', style='italic')
    
    # --- INPUT box ---
    box1 = FancyBboxPatch((1.0, 3.3), 2.2, 1.2, boxstyle="round,pad=0.15",
                          facecolor=WHITE, edgecolor=DARK, linewidth=1.5, zorder=2)
    ax.add_patch(box1)
    ax.text(2.1, 4.15, 'Input', fontsize=11, fontweight='bold', color=DARK, ha='center')
    ax.text(2.1, 3.85, 'Template + Data', fontsize=9, color='#555', ha='center')
    ax.text(2.1, 3.6, '.xlsx', fontsize=8, color='#999', ha='center')
    
    # --- SCANNER ---
    arrow1 = FancyArrowPatch((3.3, 3.9), (4.2, 3.9), arrowstyle='->', color=GRAY,
                             linewidth=1.5, mutation_scale=15, zorder=1)
    ax.add_patch(arrow1)
    box2 = FancyBboxPatch((4.3, 3.3), 2.2, 1.2, boxstyle="round,pad=0.15",
                          facecolor=WHITE, edgecolor=ORANGE, linewidth=1.5, zorder=2)
    ax.add_patch(box2)
    ax.text(5.4, 4.15, 'Scanner', fontsize=11, fontweight='bold', color=ORANGE, ha='center')
    ax.text(5.4, 3.85, 'Detect template type', fontsize=8.5, color='#555', ha='center')
    ax.text(5.4, 3.6, 'Has images? Charts?', fontsize=8, color='#999', ha='center')
    
    # --- BRANCHING ---
    # Left branch -> openpyxl
    arrow_l = FancyArrowPatch((6.6, 4.2), (7.4, 4.8), arrowstyle='->', color=GRAY,
                              linewidth=1.5, mutation_scale=15, connectionstyle="arc3,rad=0.3", zorder=1)
    ax.add_patch(arrow_l)
    # Right branch -> ZIP
    arrow_r = FancyArrowPatch((6.6, 3.6), (7.4, 3.0), arrowstyle='->', color=GRAY,
                              linewidth=1.5, mutation_scale=15, connectionstyle="arc3,rad=-0.3", zorder=1)
    ax.add_patch(arrow_r)
    
    # Left engine: openpyxl
    box3 = FancyBboxPatch((7.5, 4.3), 2.0, 1.0, boxstyle="round,pad=0.12",
                          facecolor=LIGHT_BLUE, edgecolor=BLUE, linewidth=1.5, zorder=2)
    ax.add_patch(box3)
    ax.text(8.5, 5.05, 'openpyxl Engine', fontsize=10, fontweight='bold', color=DARK, ha='center')
    ax.text(8.5, 4.75, 'Fast, no images', fontsize=8, color='#555', ha='center')
    ax.text(8.5, 4.55, 'Formulas + formatting', fontsize=7.5, color='#999', ha='center')
    
    # Right engine: ZIP
    box4 = FancyBboxPatch((7.5, 2.5), 2.0, 1.0, boxstyle="round,pad=0.12",
                          facecolor='#E8F5E9', edgecolor=GREEN, linewidth=1.5, zorder=2)
    ax.add_patch(box4)
    ax.text(8.5, 3.25, 'ZIP Engine', fontsize=10, fontweight='bold', color=GREEN, ha='center')
    ax.text(8.5, 2.95, 'Images + print settings', fontsize=8, color='#555', ha='center')
    ax.text(8.5, 2.75, 'Perfect fidelity', fontsize=7.5, color='#999', ha='center')
    
    # --- Output arrows ---
    arrow_lo = FancyArrowPatch((9.6, 4.8), (10.4, 3.9), arrowstyle='->', color=GRAY,
                               linewidth=1.5, mutation_scale=15, connectionstyle="arc3,rad=-0.3", zorder=1)
    ax.add_patch(arrow_lo)
    arrow_ro = FancyArrowPatch((9.6, 3.0), (10.4, 3.9), arrowstyle='->', color=GRAY,
                               linewidth=1.5, mutation_scale=15, connectionstyle="arc3,rad=0.3", zorder=1)
    ax.add_patch(arrow_ro)
    
    # Output box
    box5 = FancyBboxPatch((10.5, 3.3), 1.2, 1.2, boxstyle="round,pad=0.15",
                          facecolor=DARK, edgecolor=DARK, linewidth=1.5, zorder=2)
    ax.add_patch(box5)
    ax.text(11.1, 4.05, 'Output', fontsize=10, fontweight='bold', color=WHITE, ha='center')
    ax.text(11.1, 3.8, 'Filled .xlsx', fontsize=8, color='#AAC4DF', ha='center')
    ax.text(11.1, 3.55, 'or .pdf', fontsize=8, color='#AAC4DF', ha='center')
    
    # --- Feature checklist bottom ---
    features = [
        ('Auto-detect engine', 'Zero config needed'),
        ('Placeholder syntax', '{{column_name}} in cells'),
        ('Batch generation', '1 template × N rows'),
        ('PDF export', 'COM-based, full fidelity'),
        ('BQ page merge', 'Multi-sheet workbook'),
    ]
    
    for i, (title, desc) in enumerate(features):
        x = 1.2 + i * 2.2
        # Checkmark
        ax.text(x, 1.8, '✓', fontsize=13, color=GREEN, fontweight='bold', ha='left')
        ax.text(x + 0.3, 1.8, title, fontsize=8.5, fontweight='bold', color=DARK, ha='left', va='center')
        ax.text(x + 0.3, 1.55, desc, fontsize=7, color='#888', ha='left', va='center')
    
    # Tagline
    ax.text(0.5, 0.6, 'pip install -r requirements.txt  →  python example_basic.py  →  Done.',
           fontsize=9, color='#aaa', style='italic', ha='left')
    
    plt.tight_layout(pad=0.5)
    fig.savefig(out_path, dpi=150, bbox_inches='tight', facecolor=BG, edgecolor='none')
    plt.close(fig)
    print(f"✓ Saved: {out_path} ({Path(out_path).stat().st_size // 1024} KB)")

if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "demo/demo_preview.png"
    render(out)
