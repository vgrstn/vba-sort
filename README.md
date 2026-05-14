# vba-sort
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

Generic VBA standard module for sorting, searching, and shuffling one-dimensional arrays.

Uses a **hybrid QuickSort + InsertionSort** algorithm (threshold 47, as used by Java's `Arrays.sort`). The sort is iterative — no recursion, no stack overflow risk. Random pivot selection avoids worst-case behaviour on already-sorted input.

---

## 📦 Features

- **In-place or by-index** — sort and search either modify the array directly or operate through a `Long()` index array, leaving the original untouched
- **Ascending or descending** — optional `asc` parameter reverses the sort order
- **Binary search** — O(log n) search on a sorted array; returns the lowest index for duplicate values
- **Consecutive search** — find the next occurrence of a value after a prior result
- **Shuffle** — Fisher-Yates shuffle
- Works with any comparable VBA type (numbers, strings, dates, Variants)
- Pure VBA, zero dependencies, Rubberduck-friendly annotations

---

## ⚙️ Public Interface

| Member | Description |
|---|---|
| `SortArray arr [, idx [, asc]]` | Sorts `arr` in place, or by index if `idx` is provided. `asc = False` reverses the order. |
| `SearchArray(arr, value [, idx [, start]])` | Binary search in a sorted array. Returns the lowest matching index, or `Null` if not found. Pass a prior result as `start` to find the next duplicate. |
| `IsArraySorted(arr [, idx])` | Returns `True` if `arr` is sorted (ascending or descending). |
| `ShuffleArray arr` | Randomises the order of elements in `arr` (Fisher-Yates). |

All methods require a non-empty one-dimensional array. `vbErrorTypeMismatch (13)` is raised for multi-dimensional or empty arrays.

---

## 🚀 Quick Start

```vb
' Sort in place
Dim a() As Variant
a = Array(5, 3, 8, 1, 4)
SortArray a                          ' -> (1, 3, 4, 5, 8)

' Sort descending
SortArray a, asc:=False              ' -> (8, 5, 4, 3, 1)

' Sort by index (original array unchanged)
Dim idx As Variant
SortArray a, idx                     ' idx holds sorted positions

' Binary search
SortArray a
Dim pos As Variant
pos = SearchArray(a, 4)              ' -> index of 4, or Null

' Find all duplicates
a = Array(1, 2, 2, 2, 3)
SortArray a
pos = SearchArray(a, 2)              ' first occurrence
Do While Not IsNull(pos)
    Debug.Print pos                  ' prints each index where a(i) = 2
    pos = SearchArray(a, 2, , pos)   ' next occurrence
Loop

' Shuffle
ShuffleArray a

' Check if sorted
Debug.Print IsArraySorted(a)         ' True or False
```

---

## 🔑 Index-based sorting

When `idx` is passed to `SortArray`, the original array is left unchanged and a `Long()` index array is returned. All other methods accept the same `idx` to operate in index space:

```vb
Dim names() As Variant
names = Array("Charlie", "Alice", "Bob")

Dim idx As Variant
SortArray names, idx
' names is unchanged: ("Charlie", "Alice", "Bob")
' idx maps sorted order: names(idx(0)) = "Alice", etc.

Dim pos As Variant
pos = SearchArray(names, "Bob", idx)
Debug.Print IsArraySorted(names, idx)  ' -> True
```

---

## 🧠 Algorithm

| Phase | Algorithm | Condition |
|---|---|---|
| Large partitions | Randomised QuickSort | partition size ≥ 47 |
| Small partitions | Insertion sort | partition size < 47 |

- **Iterative** — uses a `Collection` as an explicit stack; no recursion
- **Smaller-first** — always processes the smaller partition immediately, pushes the larger one; keeps stack depth at O(log n)
- **Random pivot** — guards against worst-case O(n²) on sorted or reverse-sorted input
- **Direction-aware binary search** — detects ascending vs. descending order from first and last elements; works correctly for both

---

## 📄 License

MIT © 2025 Vincent van Geerestein
