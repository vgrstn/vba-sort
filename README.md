# vba-sort
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

Generic VBA standard module for sorting, searching, and shuffling one-dimensional arrays.

Uses a **hybrid QuickSort + InsertionSort** algorithm (threshold 47, as used by Java's `Arrays.sort`). The sort is iterative — no recursion, no stack overflow risk. Random pivot selection avoids worst-case behaviour on already-sorted input.

---

## 📁 Files

| File | Description |
|---|---|
| `Sort.bas` | Source file with [Rubberduck](https://rubberduckvba.com/) annotations (`'@Description`, `'@IgnoreModule`) |
| `Sort_WithAttributes.bas` | Ready-to-import version with VB attributes baked in — no Rubberduck required |

Both files are identical in behaviour. Import `Sort_WithAttributes.bas` if you are not using Rubberduck.

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

**Return values in index mode:** `SearchArray` returns a position within `idx`, not within `arr`. To retrieve the element, use `arr(idx(pos))`. This is consistent with `SortArray` and `IsArraySorted`, which also operate entirely in `idx` space.

---

## 🔍 Search return values

`SearchArray` always returns the **lowest** index at which the value occurs. For duplicate values, use the `start` parameter to iterate:

```vb
pos = SearchArray(arr, value)           ' first occurrence (index into arr, or into idx)
pos = SearchArray(arr, value, , pos)    ' next occurrence
pos = SearchArray(arr, value, , pos)    ' next occurrence, etc.
```

Returns `Null` when no further match is found. Because the array is sorted, only the single adjacent position needs to be checked on each continuation call — O(1) per step.

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

## 🧠 Implementation notes

### Iterative QuickSort with Collection-as-stack

The sort uses an explicit `Collection` instead of recursion. The collection is pre-seeded with one `Empty` sentinel item so that the empty-stack condition is simply `Stack.Count = 1`, without needing a separate flag. Items are inserted and removed at position 1, giving LIFO order.

### Smaller-partition-first

After each partition, the smaller half is processed immediately in the current loop iteration and the larger half is pushed onto the stack. This limits stack depth to O(log n) regardless of pivot quality — the stack never holds more than log₂(n) entries.

### Direction-aware binary search

The sort direction is inferred from the first and last elements at search time:

```vb
Dim Order As Long: Order = Compare(arr(lower), arr(upper))
```

`Order` is −1 for ascending and +1 for descending. A single `Case Order` branch then correctly narrows the search window for both directions without any `asc` parameter — the same code path handles both.

### `IsArraySorted` direction inference

The same `Order` trick is used: `Compare(Current, Previous) = Order` detects a violation (a step in the wrong direction). Equal adjacent elements are always permitted (`Compare` returns 0, which ≠ Order). A single-element array always returns `True`.

### InsertionSort is stable

Below the threshold of 47 elements, InsertionSort is used. InsertionSort is stable — equal elements keep their original relative order. QuickSort is not stable, so the overall sort is stable only within the small-partition regions.

### Fisher-Yates shuffle

```vb
For i = UBound(arr) To lower + 1 Step -1
    index = lower + Int((i - lower + 1) * Rnd)
    x = arr(i): arr(i) = arr(index): arr(index) = x
Next
```

The Durstenfeld variant: at step `i`, a random position in `[lower, i]` is chosen and swapped with `i`. Every permutation is equally likely, O(n). The `lower +` offset handles arbitrary array bases correctly.

### `IsIndexArray` validation

Before any by-index operation, the index array is validated in O(n): same length as `arr`, all values within `[LBound(arr), UBound(arr)]`, each used exactly once (checked via a Boolean bitmap). An invalid index array raises `vbErrorInvalidProcedureCall (5)`.

---

## ⚠️ Known behaviour

### String arrays sort lexicographically

`Compare` uses VBA's native `<` and `=` operators on Variants. For string arrays this means lexicographic order: `"10" < "9"` is `True`. If numeric ordering of string-encoded numbers is needed, convert to numbers before sorting.

### Mixed types raise a runtime error

Comparing incompatible Variant types (e.g. a `String` against a `Date`) raises `vbErrorTypeMismatch (13)` inside `Compare`. All elements in the array must be of mutually comparable types.

### Null elements sort to the end of an ascending array

VBA comparison operators return `Null` (not `True` or `False`) when either operand is `Null`. `Compare` treats this as "not less than, not equal" and returns 1, making `Null` appear greater than every other value. Null elements therefore accumulate at the end of an ascending sort, but their relative order among themselves is unspecified. Descending sorts with Null elements produce unreliable results.

### Duplicate-heavy arrays may be slower

The partition scheme does not implement three-way partitioning (Dutch National Flag). Elements equal to the pivot can land on either side. On arrays where most elements are identical this can degrade toward O(n²). For data with many duplicates consider pre-filtering duplicates or using a different approach.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
