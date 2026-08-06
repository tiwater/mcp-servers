# PPTX rendered-finding mapping contract review

## Stable port

| Contract item | Decision |
| --- | --- |
| capability | `tiwater.pptx-render-finding-map/v1` |
| owner | `tiwater-pptx` owns mapping a hash-bound WPS raster finding to current PPTX slide/layout/master objects. WPS/`tiwater-convert` owns the raster. |
| primary input | Current `inspect --detail` readback, a complete native-render manifest whose artifact/page hashes bind current PPTX bytes and PNGs, plus raster findings carrying page, raster hash, pixel region, kind, and optional observed text. |
| machine output | A closed finding map with every intersecting current object, `unique`/`ambiguous`/`unmapped` status, stable object locator and explicit no-operation disposition; an independently recomputed verdict. |
| non-goals | Detecting business meaning, inventing raster findings, OCR, moving an object to an inferred aesthetic target, slide-wide scaling, changing master/layout choice, or treating a unique object binding as permission to edit it. |

The producer maps EMU bounds to the exact raster dimensions and considers the
active slide, layout and master independently. Optional observed text filters
only current object text; it cannot create a target. The independent validator
recomputes candidates without calling the producer decision function.

Object identity is not edit authority. Occlusion, text overflow, picture
clipping, master/layout objects, multiple candidates and unmapped pixels all
produce no operation until a separate correction contract can prove a visible
improvement without moving, scaling, or otherwise changing unrelated content.

## Frozen semantic-variation cases

| Case | Variation | Expected result |
| --- | --- | --- |
| unseen slide text | arbitrary text, shape id and coordinates | unique slide binding; no correction inferred from identity alone |
| master footer | different layer and text | unique object identity, no automatic geometry operation |
| layout picture | no text, different shape kind | unique object identity, no automatic geometry operation |
| group + shape overlap | different kinds and overlapping regions | ambiguous; no target selected |
| duplicate visible text | repeated text in distinct objects | ambiguous unless the pixel region leaves exactly one candidate |
| changed raster hash | same page number, different bytes | fail before mapping |
| forged output target | validator receives a changed shape id | independent verdict fails |

The cases contain no scenario id, work-item id, customer value, slide number,
known filename, fixed production coordinate or expected customer answer.
