# OPC HDA Read Processed Guide

This guide explains **Read Processed** for the remote Matrikon OPC HDA client script:

- Script: `opc_hda_readprocessed.py`
- Server: `Matrikon.OPC.Simulation.1`
- Client Automation: `Matrikon.OPCHDA.Automation.1`
- Typical tag: `Random.Int1`

Read Processed asks the HDA server to return values on a **fixed time grid** (resample interval), using an **aggregate** (how each grid point is calculated from raw history).

---

## Read Raw vs Read Processed

| | Read Raw | Read Processed |
|--|----------|----------------|
| Times | Whenever raw samples exist | Fixed spacing (e.g. every 10 seconds) |
| Values | Actual stored samples | Calculated / estimated from raw samples |
| Extra inputs | Max values, bounds | **Resample interval** + **Aggregate** |

Example resample interval in Matrikon HDA Explorer: **10 Seconds**.

---

## What is Interpolative (Interpolated)?

**Interpolative** means: for each fixed clock time on the grid, the server **estimates** the tag value by drawing a straight line between the nearest raw samples **before** and **after** that time.

It does **not** invent a new plant measurement from nowhere. It answers:

> “At exactly this timestamp, what would the value be between the raw points we already have?”

### Simple example

Raw history (uneven times):

| Time | Value |
|------|-------|
| 11:00:00 | 10 |
| 11:00:10 | 30 |

Read Processed settings:

- Aggregate = **INTERPOLATIVE**
- Resample interval = **5 seconds**

Server returns a regular 5s grid:

| Time | Value | Why |
|------|-------|-----|
| 11:00:00 | 10 | at / near raw point |
| 11:00:05 | **20** | halfway between 10 and 30 (linear interpolate) |
| 11:00:10 | 30 | at / near raw point |

So at `11:00:05` there was **no raw sample**; Interpolative filled **20**.

### Quality flag

In the Python client, Interpolative points often show:

```text
Good, Non-specific [Interpolated]
```

If there is not enough raw history around that tick:

```text
Bad, Non-specific [NoData]
```

---

## Other aggregates (how they calculate)

All of these use the same idea: split the time range into buckets of size **resample interval** (e.g. 10s), then compute one result per bucket.

Assume raw samples inside one 10-second bucket `11:00:00 → 11:00:10`:

| Time | Value |
|------|-------|
| 11:00:01 | 10 |
| 11:00:04 | 40 |
| 11:00:09 | 20 |

### 1. INTERPOLATIVE (`interpolative`)

- **Meaning:** Estimate value **at** each grid timestamp by linear interpolation between neighboring raw points.
- **Explorer name:** INTERPOLATIVE
- **Typical quality:** `[Interpolated]`
- **Use when:** You want a smooth regular series (trending / plotting).

### 2. TIMEAVERAGE (`timeaverage`)

- **Meaning:** **Time-weighted average** of the raw values over the resample interval.
- **Explorer name:** TIMEAVERAGE
- **Typical quality:** `[Calculated]`
- **Simple idea:** Values that last longer in the bucket count more than short spikes.
- **Use when:** You want “average behavior” over each 10s window (not just a point estimate).

> Interpolative and Timeaverage are **different math**. Same timestamp can show different numbers. That is expected.

### 3. MINIMUM (`minimum`)

- **Meaning:** Smallest raw value inside the resample interval.
- **Explorer name:** MINIMUM
- **Example (bucket above):** minimum = **10**
- **Typical quality:** `[Calculated]`

### 4. MAXIMUM (`maximum`)

- **Meaning:** Largest raw value inside the resample interval.
- **Explorer name:** MAXIMUM
- **Example (bucket above):** maximum = **40**
- **Typical quality:** `[Calculated]`

### 5. MINIMUMACTUALTIME (`minimumactualtime`)

- **Meaning:** Minimum value in the interval, tied to the **actual timestamp when that minimum occurred** (not only the grid time).
- **Explorer name:** MINIMUMACTUALTIME
- **Example:** min = 10 at raw time `11:00:01`
- **Use when:** You need “what was the low, and when did it happen?”

### 6. MAXIMUMACTUALTIME (`maximumactualtime`)

- **Meaning:** Maximum value in the interval, tied to the **actual timestamp when that maximum occurred**.
- **Explorer name:** MAXIMUMACTUALTIME
- **Example:** max = 40 at raw time `11:00:04`
- **Use when:** You need “what was the high, and when did it happen?”

---

## Aggregates supported in this script

The script accepts these names (or numeric OPC HDA IDs):

| Name to type | ID | Notes |
|--------------|----|--------|
| `interpolative` | 1 | Primary / Explorer INTERPOLATIVE |
| `timeaverage` | 4 | Explorer TIMEAVERAGE |
| `minimumactualtime` | 7 | Explorer MINIMUMACTUALTIME |
| `minimum` | 8 | Explorer MINIMUM |
| `maximumactualtime` | 9 | Explorer MAXIMUMACTUALTIME |
| `maximum` | 10 | Explorer MAXIMUM |

Other OPC HDA aggregate names also exist in the script map, but **Matrikon Simulation may only support the ones shown in HDA Explorer**. Prefer the six above for testing.

---

## Requirements

- Windows PC2 client
- **Python 3.11 32-bit** (use `env32`)
- `pywin32`
- Network access + credentials to remote OPC PC
- Matrikon OPC Simulation HDA running on server PC

---

## How to run

Activate the 32-bit venv, then from the project folder:

```powershell
cd C:\Users\saisu\OneDrive\Desktop\opcdaresearch
.\env32\Scripts\Activate.ps1
```

### Interactive (prompts for time range, interval, aggregate)

```powershell
python opc_hda_readprocessed.py --host 192.168.1.141
```

At prompts:

- Start / End → press Enter for last 10 minutes, or type exact times
- Resample interval → e.g. `10`
- Aggregate → e.g. `interpolative` or `timeaverage`

### Non-interactive examples

**Interpolative (10 seconds):**

```powershell
python opc_hda_readprocessed.py --host 192.168.1.141 --interval 10 --aggregate interpolative --no-prompt
```

**Time average:**

```powershell
python opc_hda_readprocessed.py --host 192.168.1.141 --interval 10 --aggregate timeaverage --no-prompt
```

**Minimum:**

```powershell
python opc_hda_readprocessed.py --host 192.168.1.141 --interval 10 --aggregate minimum --no-prompt
```

**Maximum:**

```powershell
python opc_hda_readprocessed.py --host 192.168.1.141 --interval 10 --aggregate maximum --no-prompt
```

**Minimum actual time:**

```powershell
python opc_hda_readprocessed.py --host 192.168.1.141 --interval 10 --aggregate minimumactualtime --no-prompt
```

**Maximum actual time:**

```powershell
python opc_hda_readprocessed.py --host 192.168.1.141 --interval 10 --aggregate maximumactualtime --no-prompt
```

### Useful optional arguments

```powershell
python opc_hda_readprocessed.py --host 192.168.1.141 --tag Random.Int1 --minutes 10 --interval 10 --aggregate interpolative
```

| Argument | Meaning | Default |
|----------|---------|---------|
| `--host` | OPC server IP / hostname | `192.168.1.141` |
| `--tag` | Item ID | `Random.Int1` |
| `--minutes` | Default history window if using defaults / `--no-prompt` | `10` |
| `--interval` | Resample interval seconds | `10` |
| `--aggregate` | Aggregate name or ID | `interpolative` |
| `--no-prompt` | Skip interactive prompts | off |
| `--user` / `--domain` / `--password` | Windows credentials | script defaults |

---

## How to compare with Matrikon HDA Explorer

For a fair match, set **exactly the same** in both tools:

1. Same server / tag (`Random.Int1`)
2. Same **Start** and **End** (same seconds)
3. Same **Resample Interval** (e.g. 10 Seconds)
4. Same **Aggregate** (Interpolative vs Timeaverage are different)

### Timezone note

- Explorer usually shows **local time** (e.g. IST)
- This script prints timestamps in **UTC**

Conversion for India (IST = UTC+5:30):

```text
06:06:35 UTC  =  11:36:35 IST
```

So the clocks can match even when the printed strings look different.

### Why values might still differ

- Different aggregate (Interpolative ≠ Timeaverage)
- Different start/end (bucket alignment changes)
- `Random.Int1` history is live / rolling — run both at nearly the same time
- Leading `NoData` rows mean that part of the window had no usable raw history yet

---

## What the script prints

For each returned point:

| Column | Meaning |
|--------|---------|
| DATAVALUE | Processed value from server (`None` if NoData) |
| TIMESTAMP (UTC) | Server timestamp shown in UTC |
| QUALITY | Explorer-style text, e.g. `Good, Non-specific [Interpolated]` or `[Calculated]` |

No timestamp reconstruction is applied. Values/qualities/timestamps come from the HDA Automation response.

---

## Quick test checklist

1. Run Interpolative in Python and Explorer with same window → compare UTC↔IST times and values  
2. Run Timeaverage with same window → expect `[Calculated]` and different numbers than Interpolative  
3. Run Minimum / Maximum → expect lows/highs per 10s bucket  
4. Run MinimumActualTime / MaximumActualTime → confirm against Explorer  

---

## Related scripts

| Script | Purpose |
|--------|---------|
| `opc_hda_remote1.py` | Remote **ReadRaw** (UTC display) |
| `opc_hda_readprocessed.py` | Remote **ReadProcessed** (this guide) |
| `opc_hda_read_test.py` | Older ReadRaw test variant |
