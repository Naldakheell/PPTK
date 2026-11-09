# PPTK cyan ribbon integration manifest

These 16px icons are installed in `src/PPTK/Ribbon/CustomUI.xml` and in `src/PPTK-WINDOWS-FINAL.pptm`. The source ribbon and the embedded ribbon are the same file.

Apply this only after the Smart Bar, Align Matrix, and Align Central UI task has finished editing the ribbon and the PPTM.

## Brand colors

| Role | Hex | VBA RGB | OLE Long |
| --- | --- | --- | ---: |
| Primary cyan | #23BAEC | RGB(35, 186, 236) | 15514147 |
| White | #FFFFFF | RGB(255, 255, 255) | 16777215 |
| Dark text | #30343B | RGB(48, 52, 59) | 3879984 |
| Light background | #F7F9FB | RGB(247, 249, 251) | 16513527 |
| Subtle border | #E3E8ED | RGB(227, 232, 237) | 15591651 |

UserForm helpers live in `ModulePPTKBrand.bas`. That module is imported in the Windows VBA project. `PPTKSmartBarForm`, `PPTKMatrixForm`, and `PPTKCentralForm` call it. Do not import it into the Mac workbook.

| Helper | Result |
| --- | --- |
| PPTK_StyleForm | White background, dark text |
| PPTK_StylePrimaryButton | Cyan background, white text, no border |
| PPTK_StyleSecondaryButton | Light background, dark text, subtle border |
| PPTK_StyleField | White background, dark text, subtle border |

## Icon files

Embed the 16px PNGs from `icons/16`. The current ribbon uses `size="normal"`, which is a 16px image in the Office 2006 custom UI schema. The 32px PNGs in `icons/32` are the same artwork at 2x for inspection or a later large-button test. Do not embed a 32px file on a normal button; Office will scale it down.

Backgrounds are transparent. Cyan is #23BAEC. Ink is #30343B. White is used only as a cutout.

## Control mapping

Edit each control by its `id`. Do not search-and-replace an `imageMso` value: several controls share one Office icon today and need different cyan icons.

`replace` removes that control's `imageMso` attribute and adds `image`. `add` only adds `image` on a custom control that currently has no icon. Leave `id`, `label`, `onAction`, `tag`, `size`, `screentip`, `supertip`, and group order unchanged.

| Group | Control id | Action | Remove imageMso | image value | 16px file |
| --- | --- | --- | --- | --- | --- |
| Design | `DesignMoreFormatting` | replace | ShapeFillColorPicker | `pptk-format-more` | icons/16/pptk-format-more.png |
| Design | `PPTKFixSlideColors` | replace | ShapeFillColorPicker | `pptk-fix-colors` | icons/16/pptk-fix-colors.png |
| Design | `PPTKCopyFormatting` | replace | FormatPainter | `pptk-copy-format` | icons/16/pptk-copy-format.png |
| Design | `PPTKPasteFormatting` | replace | FormatPainter | `pptk-paste-format` | icons/16/pptk-paste-format.png |
| Reuse | `SaveSelectedSlidesButton` | replace | FileSaveAs | `pptk-save-slides` | icons/16/pptk-save-slides.png |
| Reuse | `EmailSelectedSlidesButton` | add |  | `pptk-email-slides` | icons/16/pptk-email-slides.png |
| Reuse | `SaveToLibraryButton` | add |  | `pptk-library-save` | icons/16/pptk-library-save.png |
| Reuse | `OpenSlideLibraryButton` | add |  | `pptk-library-open` | icons/16/pptk-library-open.png |
| Insert | `PPTKInsertT` | replace | TextBoxInsert | `pptk-text` | icons/16/pptk-text.png |
| Insert | `PPTKInsertN` | replace | SlideNew | `pptk-new-slide` | icons/16/pptk-new-slide.png |
| Insert | `PPTKInsertI` | replace | IconInsertFromFile | `pptk-insert-icon` | icons/16/pptk-insert-icon.png |
| Insert | `PPTKEncloseIcons` | replace | ShapeOval | `pptk-enclose` | icons/16/pptk-enclose.png |
| Insert | `InsertMoreMenu` | replace | PasteAlternative | `pptk-insert-more` | icons/16/pptk-insert-more.png |
| Insert | `PPTKAddRightTK` | replace | TextBoxInsert | `pptk-tk-right` | icons/16/pptk-tk-right.png |
| Insert | `PPTKAddLeftTK` | replace | TextBoxInsert | `pptk-tk-left` | icons/16/pptk-tk-left.png |
| Insert | `PPTKAddTopTK` | replace | TextBoxInsert | `pptk-tk-top` | icons/16/pptk-tk-top.png |
| Insert | `PPTKAddBottomTK` | replace | TextBoxInsert | `pptk-tk-bottom` | icons/16/pptk-tk-bottom.png |
| Text | `TextMoreMenu` | replace | MenuView2 | `pptk-text-more` | icons/16/pptk-text-more.png |
| Text | `SpecialCharacterMenu` | replace | SymbolInsert | `pptk-special-char` | icons/16/pptk-special-char.png |
| Text | `PPTKPasteContentOnly` | replace | PasteTextOnly | `pptk-paste-text` | icons/16/pptk-paste-text.png |
| Text | `PPTKMergeText` | replace | TextBoxInsert | `pptk-merge-text` | icons/16/pptk-merge-text.png |
| Text | `PPTKRemoveMargins` | replace | ParagraphMarks | `pptk-remove-margins` | icons/16/pptk-remove-margins.png |
| Align | `ObjectsAlignLeftsButton` | replace | ObjectsAlignLeft | `pptk-align-left` | icons/16/pptk-align-left.png |
| Align | `ObjectsAlignCentersButton` | replace | ObjectsAlignCenterHorizontal | `pptk-align-center` | icons/16/pptk-align-center.png |
| Align | `ObjectsAlignRightsButton` | replace | ObjectsAlignRight | `pptk-align-right` | icons/16/pptk-align-right.png |
| Align | `ObjectsAlignTopsButton` | replace | ObjectsAlignTop | `pptk-align-top` | icons/16/pptk-align-top.png |
| Align | `ObjectsAlignMiddlesButton` | replace | ObjectsAlignMiddleVertical | `pptk-align-middle` | icons/16/pptk-align-middle.png |
| Align | `ObjectsAlignBottomsButton` | replace | ObjectsAlignBottom | `pptk-align-bottom` | icons/16/pptk-align-bottom.png |
| Space | `ObjectsDistributeHorizontallyButton` | replace | AlignDistributeHorizontally | `pptk-dist-h` | icons/16/pptk-dist-h.png |
| Space | `ObjectsDistributeVerticallyButton` | replace | AlignDistributeVertically | `pptk-dist-v` | icons/16/pptk-dist-v.png |
| Space | `ObjectsResizeAndSpaceEvenly` | replace | AlignRelativeToPage | `pptk-resize-space` | icons/16/pptk-resize-space.png |
| Space | `ResizeAndSpaceEvenHorizontal` | add |  | `pptk-resize-h` | icons/16/pptk-resize-h.png |
| Space | `ResizeAndSpaceEvenHorizontalPreserveFirst` | add |  | `pptk-resize-h-first` | icons/16/pptk-resize-h-first.png |
| Space | `ResizeAndSpaceEvenHorizontalPreserveLast` | add |  | `pptk-resize-h-last` | icons/16/pptk-resize-h-last.png |
| Space | `ResizeAndSpaceEvenVertical` | add |  | `pptk-resize-v` | icons/16/pptk-resize-v.png |
| Space | `ResizeAndSpaceEvenVerticalPreserveFirst` | add |  | `pptk-resize-v-first` | icons/16/pptk-resize-v-first.png |
| Space | `ResizeAndSpaceEvenVerticalPreserveLast` | add |  | `pptk-resize-v-last` | icons/16/pptk-resize-v-last.png |
| Space | `SpaceGapsMenu` | replace | AlignDistributeHorizontally | `pptk-gaps` | icons/16/pptk-gaps.png |
| Space | `ObjectsRemoveSpaceHorizontal` | replace | ObjectsAlignDistributeHorizontallyRemove | `pptk-gap-h-left` | icons/16/pptk-gap-h-left.png |
| Space | `ObjectsRemoveSpaceHorizontalRight` | replace | ObjectsAlignDistributeHorizontallyRemove | `pptk-gap-h-right` | icons/16/pptk-gap-h-right.png |
| Space | `ObjectsIncreaseSpacingHorizontal` | replace | HorizontalSpacingIncrease | `pptk-gap-h-inc` | icons/16/pptk-gap-h-inc.png |
| Space | `ObjectsDecreaseSpacingHorizontal` | replace | HorizontalSpacingDecrease | `pptk-gap-h-dec` | icons/16/pptk-gap-h-dec.png |
| Space | `ObjectsRemoveSpaceVertical` | replace | ObjectsAlignDistributeVerticallyRemove | `pptk-gap-v-up` | icons/16/pptk-gap-v-up.png |
| Space | `ObjectsRemoveSpaceVerticalDown` | replace | ObjectsAlignDistributeVerticallyRemove | `pptk-gap-v-down` | icons/16/pptk-gap-v-down.png |
| Space | `ObjectsIncreaseSpacingVertical` | replace | VerticalSpacingIncrease | `pptk-gap-v-inc` | icons/16/pptk-gap-v-inc.png |
| Space | `ObjectsDecreaseSpacingVertical` | replace | VerticalSpacingDecrease | `pptk-gap-v-dec` | icons/16/pptk-gap-v-dec.png |
| Shape | `PPTKChangeShapeMenu` | replace | ShapeChangeShapeGallery | `pptk-change-shape` | icons/16/pptk-change-shape.png |
| Shape | `ObjectsSameWidthButton` | replace | ColumnWidth | `pptk-same-width` | icons/16/pptk-same-width.png |
| Shape | `ObjectsSameHeightButton` | replace | RowHeight | `pptk-same-height` | icons/16/pptk-same-height.png |
| Shape | `ObjectsSameHeightAndWidthButton` | replace | SizeToControlHeightAndWidth | `pptk-same-size` | icons/16/pptk-same-size.png |
| Shape | `ObjectsSwapPosition` | replace | CircularReferences | `pptk-swap` | icons/16/pptk-swap.png |
| Shape | `ObjectsSwapPositionMenu` | replace | CircularReferences | `pptk-swap` | icons/16/pptk-swap.png |
| Shape | `ObjectsSwapPositionCentered` | replace | CircularReferences | `pptk-swap-center` | icons/16/pptk-swap-center.png |
| Shape | `PPTKSwapText` | replace | ReplaceDialog | `pptk-swap-text` | icons/16/pptk-swap-text.png |
| Shape | `ShapeMoreMenu` | replace | ShapesMoreShapes | `pptk-shape-more` | icons/16/pptk-shape-more.png |
| Shape | `ObjectsSizeToTallest` | replace | RowHeight | `pptk-size-tall` | icons/16/pptk-size-tall.png |
| Shape | `ObjectsSizeShortest` | replace | RowHeight | `pptk-size-short` | icons/16/pptk-size-short.png |
| Shape | `ObjectsSizeToWidest` | replace | ColumnWidth | `pptk-size-wide` | icons/16/pptk-size-wide.png |
| Shape | `ObjectsSizeNarrowest` | replace | ColumnWidth | `pptk-size-narrow` | icons/16/pptk-size-narrow.png |
| Shape | `PPTKAddRightColumn` | replace | TableInsert | `pptk-col-add` | icons/16/pptk-col-add.png |
| Shape | `PPTKAddBottomRow` | replace | TableInsert | `pptk-row-add` | icons/16/pptk-row-add.png |
| Shape | `PPTKDeleteRightColumn` | replace | TableInsert | `pptk-col-del` | icons/16/pptk-col-del.png |
| Shape | `PPTKDeleteBottomRow` | replace | TableInsert | `pptk-row-del` | icons/16/pptk-row-del.png |
| Smart Align | `ObjectsArrangeShapesButton` | replace | SlidesPerPage9Slides | `pptk-arrange` | icons/16/pptk-arrange.png |
| Smart Align | `ObjectsStretchMenu` | replace | DrawingCanvasExpand | `pptk-stretch` | icons/16/pptk-stretch.png |
| Smart Align | `ObjectsStretchLeft` | add |  | `pptk-stretch-left` | icons/16/pptk-stretch-left.png |
| Smart Align | `ObjectsStretchRight` | add |  | `pptk-stretch-right` | icons/16/pptk-stretch-right.png |
| Smart Align | `ObjectsStretchTop` | add |  | `pptk-stretch-top` | icons/16/pptk-stretch-top.png |
| Smart Align | `ObjectsStretchBottom` | add |  | `pptk-stretch-bottom` | icons/16/pptk-stretch-bottom.png |
| Smart Align | `ObjectsStretchLeftShapeRight` | add |  | `pptk-stretch-to-left` | icons/16/pptk-stretch-to-left.png |
| Smart Align | `ObjectsStretchRightShapeLeft` | add |  | `pptk-stretch-to-right` | icons/16/pptk-stretch-to-right.png |
| Smart Align | `ObjectsStretchTopShapeBottom` | add |  | `pptk-stretch-to-top` | icons/16/pptk-stretch-to-top.png |
| Smart Align | `ObjectsStretchBottomShapeTop` | add |  | `pptk-stretch-to-bottom` | icons/16/pptk-stretch-to-bottom.png |
| Smart Align | `SmartAlignMoreMenu` | replace | GroupPosition | `pptk-align-more` | icons/16/pptk-align-more.png |
| Smart Align | `PPTKMatchCornerRoundness` | add |  | `pptk-corner` | icons/16/pptk-corner.png |
| Shape | `PPTKUnifyShapeType` | add |  | `pptk-unify` | icons/16/pptk-unify.png |
| Shape | `PPTKGridMenu` | add |  | `pptk-matrix` | icons/16/pptk-matrix.png |
| Shape | `PPTKAddRightColumnExpand` | add |  | `pptk-col-expand` | icons/16/pptk-col-expand.png |
| Shape | `PPTKAddBottomRowExpand` | add |  | `pptk-row-expand` | icons/16/pptk-row-expand.png |
| Shape | `PPTKDeleteRightColumnShrink` | add |  | `pptk-col-shrink` | icons/16/pptk-col-shrink.png |
| Shape | `PPTKDeleteBottomRowShrink` | add |  | `pptk-row-shrink` | icons/16/pptk-row-shrink.png |
| Shape | `PPTKFlipCurrentSlide` | add |  | `pptk-flip-slide` | icons/16/pptk-flip-slide.png |
| Shape | `PPTKFlipEntirePresentation` | add |  | `pptk-flip-deck` | icons/16/pptk-flip-deck.png |
| Smart Align | `PPTKAlignProcess` | replace | ShapeChevron | `pptk-process` | icons/16/pptk-process.png |
| Smart Align | `PPTKAlignMatrix` | replace | TableInsert | `pptk-matrix` | icons/16/pptk-matrix.png |
| Smart Align | `PPTKAlignCentral` | replace | GroupPosition | `pptk-central` | icons/16/pptk-central.png |
| PPTK | `PPTKShortcutR` | replace | ShapeRectangle | `pptk-shape-rect` | icons/16/pptk-shape-rect.png |
| PPTK | `PPTKShortcutU` | replace | ShapeRoundedRectangle | `pptk-shape-round` | icons/16/pptk-shape-round.png |
| PPTK | `PPTKShortcutO` | replace | ShapeOval | `pptk-shape-oval` | icons/16/pptk-shape-oval.png |
| PPTK | `PPTKShortcutC` | replace | ShapeChevron | `pptk-shape-chevron` | icons/16/pptk-shape-chevron.png |
| PPTK | `PPTKShortcutP` | replace | ShapePentagon | `pptk-shape-pentagon` | icons/16/pptk-shape-pentagon.png |
| PPTK | `PPTKShortcutG` | replace | ShapeIsoscelesTriangle | `pptk-shape-triangle` | icons/16/pptk-shape-triangle.png |
| PPTK | `PPTKShortcutM` | replace | ShapeParallelogram | `pptk-shape-para` | icons/16/pptk-shape-para.png |
| PPTK | `PPTKSmartBar` | replace | GroupPosition | `pptk-smart-bar` | icons/16/pptk-smart-bar.png |
| Tools | `ToolsTableMenu` | replace | TableInsert | `pptk-table` | icons/16/pptk-table.png |
| Tools | `ToolsMoreMenu` | replace | AdvancedFileProperties | `pptk-tools` | icons/16/pptk-tools.png |
| Tools | `CopyStorylineMenu` | replace | SpeakerNotes | `pptk-storyline` | icons/16/pptk-storyline.png |

`ObjectsSwapPosition` and `ObjectsSwapPositionMenu` share `pptk-swap` because they are the same swap command. Every other image is used by one control.

## Controls that stay on Office icons

Any control with `idMso` keeps its built-in Office icon. The 2006 ribbon schema does not accept a custom image on an `idMso` control, and replacing those controls would change command behavior. That includes fill, outline, font, paste PNG, native align toggles, and the other built-in buttons already on the PPTK tab.

## How to integrate

1. Wait until no other task is editing `src/PPTK/Ribbon/CustomUI.xml` or `src/PPTK-WINDOWS-FINAL.pptm`.
2. Re-run `Generate-PptkRibbonIcons.ps1`. It fails if a mapped control, element type, or `imageMso` no longer matches the XML.
3. In the then-current `CustomUI.xml`, apply the table above one control at a time. Do not replace the whole file from an older copy.
4. Copy `icons/16/*.png` into the PPTM part `customUI/images/`.
5. The package relationship `PPTKRibbon` already points at `customUI/customUI.xml`. Leave `_rels/.rels` unchanged.
6. `customUI/_rels/customUI.xml.rels` does not exist in the PPTM today. If it is still missing, add `customUI.xml.rels` from this folder. If the UI task created that file, insert only Relationship rows whose Id is not already present.
7. `[Content_Types].xml` already has a PNG default. Do not add another PNG content type.
8. Import `ModulePPTKBrand.bas` with the VBA editor when a UserForm is ready to call it. Do not call it from Smart Bar, Align Matrix, Align Central, shortcut dispatch, or keyboard code as part of the icon install.
9. Reload the add-in and look at the PPTK tab at normal ribbon size. The icons are not live until this step has been done.

Office RibbonX Editor can do steps 4 and 6 if the `image` values match the relationship Ids in `customUI.xml.rels`.

