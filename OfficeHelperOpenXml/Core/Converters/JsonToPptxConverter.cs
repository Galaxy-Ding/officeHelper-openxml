using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Core.Writers;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeHelperOpenXml.Core.Converters
{
    /// <summary>
    /// Converts JSON presentation data to PowerPoint PPTX files
    /// </summary>
    public class JsonToPptxConverter
    {
        private PresentationDocument _document;
        private PresentationJsonData _presentationData;
        private Dictionary<int, SlideMasterPart> _masterSlideMap;
        private Dictionary<int, SlideLayoutPart> _layoutMap;
        private ConversionLogger _logger;
        private PackageMetadataGenerator _metadataGenerator;
        private RelationshipIdManager _relationshipIdManager;
        private SlideWriter _slideWriter;
        private LayoutRelationshipGenerator _layoutRelationshipGenerator;

        /// <summary>
        /// Initializes a new instance of JsonToPptxConverter
        /// </summary>
        public JsonToPptxConverter()
        {
            _logger = new ConversionLogger();
            _metadataGenerator = new PackageMetadataGenerator();
            _relationshipIdManager = new RelationshipIdManager();
            _slideWriter = new SlideWriter(_relationshipIdManager);
            _layoutRelationshipGenerator = new LayoutRelationshipGenerator();
        }

        /// <summary>
        /// Initializes a new instance of JsonToPptxConverter with a custom logger
        /// </summary>
        /// <param name="logger">Custom logger instance</param>
        public JsonToPptxConverter(ConversionLogger logger)
        {
            _logger = logger ?? new ConversionLogger();
            _metadataGenerator = new PackageMetadataGenerator();
            _relationshipIdManager = new RelationshipIdManager();
            _slideWriter = new SlideWriter(_relationshipIdManager);
            _layoutRelationshipGenerator = new LayoutRelationshipGenerator();
        }

        /// <summary>
        /// Loads JSON data from a file and deserializes it to PresentationJsonData.
        /// Uses streaming deserialization for better memory efficiency with large files.
        /// </summary>
        /// <param name="jsonPath">Path to the JSON file</param>
        /// <returns>Deserialized presentation data</returns>
        /// <exception cref="ConversionException">Thrown when JSON file cannot be loaded or parsed</exception>
        private PresentationJsonData LoadJson(string jsonPath)
        {
            // Check if file exists
            if (!File.Exists(jsonPath))
            {
                throw new ConversionException(
                    $"JSON file not found: {jsonPath}",
                    "File I/O"
                )
                {
                    JsonPath = jsonPath
                };
            }

            try
            {
                // Use streaming deserialization for better memory efficiency
                using (var fileStream = File.OpenRead(jsonPath))
                using (var streamReader = new StreamReader(fileStream))
                using (var jsonReader = new JsonTextReader(streamReader))
                {
                    var serializer = new JsonSerializer();
                    var presentationData = serializer.Deserialize<PresentationJsonData>(jsonReader);

                    if (presentationData == null)
                    {
                        throw new ConversionException(
                            "Failed to deserialize JSON: result is null",
                            "JSON Deserialization"
                        )
                        {
                            JsonPath = jsonPath
                        };
                    }

                    return presentationData;
                }
            }
            catch (JsonException ex)
            {
                throw new ConversionException(
                    $"JSON parsing error: {ex.Message}",
                    "JSON Deserialization",
                    ex
                )
                {
                    JsonPath = jsonPath
                };
            }
            catch (IOException ex)
            {
                throw new ConversionException(
                    $"Error reading JSON file: {ex.Message}",
                    "File I/O",
                    ex
                )
                {
                    JsonPath = jsonPath
                };
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new ConversionException(
                    $"Access denied to JSON file: {ex.Message}",
                    "File I/O",
                    ex
                )
                {
                    JsonPath = jsonPath
                };
            }
        }

        /// <summary>
        /// Main conversion method that orchestrates the entire JSON to PPTX conversion process
        /// </summary>
        /// <param name="jsonPath">Path to the input JSON file</param>
        /// <param name="outputPptxPath">Path where the output PPTX file will be saved</param>
        /// <returns>True if conversion succeeds, false otherwise</returns>
        public bool Convert(string jsonPath, string outputPptxPath)
        {
            try
            {
                // Validate input path
                if (string.IsNullOrEmpty(jsonPath))
                {
                    throw new ConversionException(
                        "JSON file path cannot be empty",
                        "Input Validation"
                    );
                }

                // Validate output path
                if (string.IsNullOrEmpty(outputPptxPath))
                {
                    throw new ConversionException(
                        "Output PPTX path cannot be empty",
                        "Input Validation"
                    );
                }

                // Ensure output directory exists
                try
                {
                    var outputDir = Path.GetDirectoryName(outputPptxPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }
                }
                catch (Exception ex)
                {
                    throw new ConversionException(
                        $"Failed to create output directory: {ex.Message}",
                        "File I/O",
                        ex
                    );
                }

                _logger.LogInfo($"Starting conversion: {jsonPath} -> {outputPptxPath}");

                // Load JSON data
                _presentationData = LoadJson(jsonPath);
                _logger.LogInfo($"Loaded JSON with {_presentationData.MasterSlides.Count} master slides and {_presentationData.ContentSlides.Count} content slides");

                // Create PowerPoint document
                try
                {
                    _document = PresentationDocument.Create(outputPptxPath, PresentationDocumentType.Presentation);
                    InitializePresentation();
                }
                catch (Exception ex)
                {
                    throw new ConversionException(
                        $"Failed to create PowerPoint document: {ex.Message}",
                        "Document Creation",
                        ex
                    );
                }

                // Initialize master slide map
                _masterSlideMap = new Dictionary<int, SlideMasterPart>();
                _layoutMap = new Dictionary<int, SlideLayoutPart>();

                // Create master slides
                try
                {
                    CreateMasterSlides();
                }
                catch (ConversionException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    throw new ConversionException(
                        $"Failed to create master slides: {ex.Message}",
                        "Master Slide Creation",
                        ex
                    );
                }

                // Create content slides
                try
                {
                    CreateContentSlides();
                }
                catch (ConversionException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    throw new ConversionException(
                        $"Failed to create content slides: {ex.Message}",
                        "Content Slide Creation",
                        ex
                    );
                }

                // Save document
                try
                {
                    _document.Save();
                    _document.Dispose();
                    _document = null;
                }
                catch (Exception ex)
                {
                    throw new ConversionException(
                        $"Failed to save PowerPoint document: {ex.Message}",
                        "Document Save",
                        ex
                    );
                }

                // Log completion statistics (optimized to calculate once)
                int masterSlideCount = _presentationData.MasterSlides?.Count ?? 0;
                int contentSlideCount = _presentationData.ContentSlides?.Count ?? 0;
                int totalShapes = (_presentationData.MasterSlides?.Sum(s => s.Shapes?.Count ?? 0) ?? 0) +
                                 (_presentationData.ContentSlides?.Sum(s => s.Shapes?.Count ?? 0) ?? 0);
                
                _logger.LogSuccess($"Conversion completed successfully");
                _logger.LogInfo($"Statistics: {masterSlideCount} master slides, {contentSlideCount} content slides, {totalShapes} shapes");
                return true;
            }
            catch (ConversionException ex)
            {
                _logger.LogError($"Conversion failed in {ex.Context}: {ex.Message}");
                if (!string.IsNullOrEmpty(ex.JsonPath))
                {
                    _logger.LogError($"JSON Path: {ex.JsonPath}");
                }
                if (ex.InnerException != null)
                {
                    _logger.LogError($"Inner Exception: {ex.InnerException.Message}");
                }
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Unexpected error during conversion: {ex.Message}");
                _logger.LogError($"Stack trace: {ex.StackTrace}");
                return false;
            }
            finally
            {
                // Ensure document is disposed
                if (_document != null)
                {
                    _document.Dispose();
                    _document = null;
                }
            }
        }

        /// <summary>
        /// Initializes the presentation with basic structure
        /// </summary>
        private void InitializePresentation()
        {
            // Reset relationship ID counter for presentation-level relationships
            _relationshipIdManager.ResetCounter();
            
            var presentationPart = _document.AddPresentationPart();
            presentationPart.Presentation = new Presentation(
                new SlideIdList(),
                new SlideSize { Cx = 12192000, Cy = 6858000 },
                new NotesSize { Cx = 6858000, Cy = 9144000 },
                new DefaultTextStyle() // Add defaultTextStyle
            );

            // Apply namespace declarations using utility class
            NamespaceDeclarationApplier.ApplyPresentationNamespaces(presentationPart.Presentation);

            // Initialize master slide list
            presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList();

            // Generate all required metadata files with sequential relationship IDs
            _metadataGenerator.GenerateAllMetadata(_document, _relationshipIdManager);
        }

        /// <summary>
        /// Creates master slides from JSON data
        /// If no master slides are defined, creates a default master slide
        /// </summary>
        private void CreateMasterSlides()
        {
            if (_presentationData.MasterSlides == null || _presentationData.MasterSlides.Count == 0)
            {
                _logger.LogInfo("No master slides defined, creating default master slide");
                CreateDefaultMasterSlide();
                return;
            }

            var presentationPart = _document.PresentationPart;
            uint masterIdCounter = 2147483648U;
            int currentSlide = 0;
            int totalSlides = _presentationData.MasterSlides.Count;

            // Create a single theme part at the presentation level (in ppt/theme/)
            // This ensures the theme is in the correct location and not duplicated
            ThemePart sharedThemePart = null;
            if (_presentationData.MasterSlides.Count > 0)
            {
                sharedThemePart = presentationPart.AddNewPart<ThemePart>(_relationshipIdManager.GetNextId());
                sharedThemePart.Theme = CreateTheme();
            }

            foreach (var masterSlideData in _presentationData.MasterSlides)
            {
                currentSlide++;
                _logger.LogProgress($"Creating master slide {masterSlideData.PageNumber}: {masterSlideData.Title}", currentSlide, totalSlides);

                // Create slide master part with sequential relationship ID
                var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>(_relationshipIdManager.GetNextId());
                
                // Reference the shared theme part - AddPart for existing parts creates relationship in master's .rels file
                slideMasterPart.AddPart(sharedThemePart, _relationshipIdManager.GetNextId());
                
                // Create slide master
                slideMasterPart.SlideMaster = CreateSlideMaster();

                // Create slide layout with sequential relationship ID
                var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>(_relationshipIdManager.GetNextId());
                slideLayoutPart.SlideLayout = CreateSlideLayout();

                // Link layout to master
                slideMasterPart.SlideMaster.SlideLayoutIdList = new SlideLayoutIdList(
                    new SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) }
                );

                // Add shapes to master slide
                if (masterSlideData.Shapes != null && masterSlideData.Shapes.Count > 0)
                {
                    AddShapesToSlide(slideMasterPart.SlideMaster.CommonSlideData.ShapeTree, masterSlideData.Shapes, slideMasterPart);
                }

                // Add master to presentation
                var slideMasterId = new SlideMasterId
                {
                    Id = masterIdCounter++,
                    RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
                };
                presentationPart.Presentation.SlideMasterIdList.Append(slideMasterId);

                // Store master slide reference
                _masterSlideMap[masterSlideData.PageNumber] = slideMasterPart;
                _layoutMap[masterSlideData.PageNumber] = slideLayoutPart;
            }

            _logger.LogSuccess($"Created {_masterSlideMap.Count} master slides");
        }

        /// <summary>
        /// Creates content slides from JSON data
        /// </summary>
        private void CreateContentSlides()
        {
            if (_presentationData.ContentSlides == null || _presentationData.ContentSlides.Count == 0)
            {
                _logger.LogWarning("No content slides to create");
                return;
            }

            var presentationPart = _document.PresentationPart;
            var slideIdList = presentationPart.Presentation.SlideIdList;
            uint slideIdCounter = 256U;

            // Sort content slides by page number to preserve order
            var sortedSlides = _presentationData.ContentSlides.OrderBy(s => s.PageNumber).ToList();
            int currentSlide = 0;
            int totalSlides = sortedSlides.Count;

            foreach (var contentSlideData in sortedSlides)
            {
                currentSlide++;
                _logger.LogProgress($"Creating content slide {contentSlideData.PageNumber}: {contentSlideData.Title}", currentSlide, totalSlides);

                // Create slide part with sequential relationship ID
                var slidePart = presentationPart.AddNewPart<SlidePart>(_relationshipIdManager.GetNextId());

                // Create slide with ColorMapOverride
                slidePart.Slide = new Slide(
                    new CommonSlideData(new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(CreateGroupTransform())
                    )),
                    new P.ColorMapOverride(new A.MasterColorMapping())
                );

                // Apply namespace declarations using utility class
                NamespaceDeclarationApplier.ApplySlideNamespaces(slidePart.Slide);

                // Link to appropriate layout (use first layout if available)
                if (_layoutMap.Count > 0)
                {
                    var layoutPart = _layoutMap.Values.First();
                    // AddPart for existing parts - use sequential relationship ID
                    slidePart.AddPart(layoutPart, _relationshipIdManager.GetNextId());
                }

                // Add shapes to content slide
                if (contentSlideData.Shapes != null && contentSlideData.Shapes.Count > 0)
                {
                    AddShapesToSlide(slidePart.Slide.CommonSlideData.ShapeTree, contentSlideData.Shapes, slidePart);
                }

                // Add slide to presentation
                var slideId = new SlideId
                {
                    Id = slideIdCounter++,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                };
                slideIdList.Append(slideId);
            }

            _logger.LogSuccess($"Created {sortedSlides.Count} content slides");
        }

        /// <summary>
        /// Creates a default master slide when none are defined in JSON
        /// This ensures the presentation has a valid structure
        /// </summary>
        private void CreateDefaultMasterSlide()
        {
            var presentationPart = _document.PresentationPart;

            // Create a single theme part at the presentation level (in ppt/theme/) with sequential ID
            var themePart = presentationPart.AddNewPart<ThemePart>(_relationshipIdManager.GetNextId());
            themePart.Theme = CreateTheme();

            // Create slide master part with sequential ID
            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>(_relationshipIdManager.GetNextId());

            // Reference the theme part - AddPart for existing parts creates relationship in master's .rels file
            slideMasterPart.AddPart(themePart, _relationshipIdManager.GetNextId());

            // Create slide master
            slideMasterPart.SlideMaster = CreateSlideMaster();

            // Create slide layout with sequential ID
            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>(_relationshipIdManager.GetNextId());
            slideLayoutPart.SlideLayout = CreateSlideLayout();

            // Link layout to master
            slideMasterPart.SlideMaster.SlideLayoutIdList = new SlideLayoutIdList(
                new SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) }
            );

            // Add master to presentation
            var slideMasterId = new SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            };
            presentationPart.Presentation.SlideMasterIdList.Append(slideMasterId);

            // Store master slide reference with page number 0 (default)
            _masterSlideMap[0] = slideMasterPart;
            _layoutMap[0] = slideLayoutPart;

            _logger.LogSuccess("Created default master slide");
        }

        /// <summary>
        /// Creates a complete transform group for group shapes
        /// </summary>
        private A.TransformGroup CreateGroupTransform()
        {
            return new A.TransformGroup(
                new A.Offset { X = 0, Y = 0 },
                new A.Extents { Cx = 0, Cy = 0 },
                new A.ChildOffset { X = 0, Y = 0 },
                new A.ChildExtents { Cx = 12192000, Cy = 6858000 }
            );
        }

        /// <summary>
        /// Creates a slide master with basic structure
        /// </summary>
        private SlideMaster CreateSlideMaster()
        {
            var slideMaster = new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(CreateGroupTransform())
                )),
                new P.ColorMap
                {
                    Background1 = A.ColorSchemeIndexValues.Light1,
                    Text1 = A.ColorSchemeIndexValues.Dark1,
                    Background2 = A.ColorSchemeIndexValues.Light2,
                    Text2 = A.ColorSchemeIndexValues.Dark2,
                    Accent1 = A.ColorSchemeIndexValues.Accent1,
                    Accent2 = A.ColorSchemeIndexValues.Accent2,
                    Accent3 = A.ColorSchemeIndexValues.Accent3,
                    Accent4 = A.ColorSchemeIndexValues.Accent4,
                    Accent5 = A.ColorSchemeIndexValues.Accent5,
                    Accent6 = A.ColorSchemeIndexValues.Accent6,
                    Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
                },
                new TextStyles( // Add required text styles
                    new TitleStyle(),
                    new BodyStyle(),
                    new OtherStyle()
                )
            );

            // Apply namespace declarations using utility class
            NamespaceDeclarationApplier.ApplySlideMasterNamespaces(slideMaster);

            return slideMaster;
        }

        /// <summary>
        /// Creates a blank slide layout with proper color map override
        /// </summary>
        private SlideLayout CreateSlideLayout()
        {
            var slideLayout = new SlideLayout(
                new CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(CreateGroupTransform())
                )),
                new P.ColorMapOverride(new A.MasterColorMapping())
            ) { Type = SlideLayoutValues.Blank };
            
            // Apply namespace declarations using utility class
            NamespaceDeclarationApplier.ApplySlideLayoutNamespaces(slideLayout);
            
            return slideLayout;
        }

        /// <summary>
        /// Creates a default theme for the presentation
        /// This is required for a valid PPTX file
        /// </summary>
        private A.Theme CreateTheme()
        {
            var theme = new A.Theme() { Name = "Office Theme" };

            // Theme elements
            var themeElements = new A.ThemeElements(
                new A.ColorScheme(
                    new A.Dark1Color(new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                    new A.Light1Color(new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new A.Dark2Color(new A.RgbColorModelHex() { Val = "44546A" }),
                    new A.Light2Color(new A.RgbColorModelHex() { Val = "E7E6E6" }),
                    new A.Accent1Color(new A.RgbColorModelHex() { Val = "4472C4" }),
                    new A.Accent2Color(new A.RgbColorModelHex() { Val = "ED7D31" }),
                    new A.Accent3Color(new A.RgbColorModelHex() { Val = "A5A5A5" }),
                    new A.Accent4Color(new A.RgbColorModelHex() { Val = "FFC000" }),
                    new A.Accent5Color(new A.RgbColorModelHex() { Val = "5B9BD5" }),
                    new A.Accent6Color(new A.RgbColorModelHex() { Val = "70AD47" }),
                    new A.Hyperlink(new A.RgbColorModelHex() { Val = "0563C1" }),
                    new A.FollowedHyperlinkColor(new A.RgbColorModelHex() { Val = "954F72" })
                ) { Name = "Office" },
                new A.FontScheme(
                    new A.MajorFont(
                        new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" },
                        new A.EastAsianFont() { Typeface = "" },
                        new A.ComplexScriptFont() { Typeface = "" }
                    ),
                    new A.MinorFont(
                        new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" },
                        new A.EastAsianFont() { Typeface = "" },
                        new A.ComplexScriptFont() { Typeface = "" }
                    )
                ) { Name = "Office" },
                new A.FormatScheme(
                    new A.FillStyleList(
                        new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 50000 },
                                    new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                                new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 37000 },
                                    new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 35000 },
                                new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 15000 },
                                    new A.SaturationModulation() { Val = 350000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }
                            ),
                            new A.LinearGradientFill() { Angle = 16200000, Scaled = true }
                        ),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(new A.SchemeColor(new A.Shade() { Val = 51000 },
                                    new A.SaturationModulation() { Val = 130000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                                new A.GradientStop(new A.SchemeColor(new A.Shade() { Val = 93000 },
                                    new A.SaturationModulation() { Val = 130000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 80000 },
                                new A.GradientStop(new A.SchemeColor(new A.Shade() { Val = 94000 },
                                    new A.SaturationModulation() { Val = 135000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }
                            ),
                            new A.LinearGradientFill() { Angle = 16200000, Scaled = false }
                        )
                    ),
                    new A.LineStyleList(
                        new A.Outline(
                            new A.SolidFill(
                                new A.SchemeColor(
                                    new A.Shade() { Val = 95000 },
                                    new A.SaturationModulation() { Val = 105000 }
                                ) { Val = A.SchemeColorValues.PhColor }
                            ),
                            new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
                        ) { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center },
                        new A.Outline(
                            new A.SolidFill(
                                new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }
                            ),
                            new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
                        ) { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center },
                        new A.Outline(
                            new A.SolidFill(
                                new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }
                            ),
                            new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
                        ) { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center }
                    ),
                    new A.EffectStyleList(
                        new A.EffectStyle(
                            new A.EffectList(
                                new A.OuterShadow(
                                    new A.RgbColorModelHex(
                                        new A.Alpha() { Val = 38000 }
                                    ) { Val = "000000" }
                                ) { BlurRadius = 40000, Distance = 20000, Direction = 5400000, RotateWithShape = false }
                            )
                        ),
                        new A.EffectStyle(
                            new A.EffectList(
                                new A.OuterShadow(
                                    new A.RgbColorModelHex(
                                        new A.Alpha() { Val = 35000 }
                                    ) { Val = "000000" }
                                ) { BlurRadius = 40000, Distance = 23000, Direction = 5400000, RotateWithShape = false }
                            )
                        ),
                        new A.EffectStyle(
                            new A.EffectList(
                                new A.OuterShadow(
                                    new A.RgbColorModelHex(
                                        new A.Alpha() { Val = 35000 }
                                    ) { Val = "000000" }
                                ) { BlurRadius = 40000, Distance = 23000, Direction = 5400000, RotateWithShape = false }
                            )
                        )
                    ),
                    new A.BackgroundFillStyleList(
                        new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 40000 },
                                    new A.SaturationModulation() { Val = 350000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                                new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 45000 },
                                    new A.Shade() { Val = 99000 },
                                    new A.SaturationModulation() { Val = 350000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 40000 },
                                new A.GradientStop(new A.SchemeColor(new A.Shade() { Val = 20000 },
                                    new A.SaturationModulation() { Val = 255000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }
                            ),
                            new A.PathGradientFill() { Path = A.PathShadeValues.Circle },
                            new A.TileRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 }
                        ) { RotateWithShape = true },
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 80000 },
                                    new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                                new A.GradientStop(new A.SchemeColor(new A.Shade() { Val = 30000 },
                                    new A.SaturationModulation() { Val = 200000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }
                            ),
                            new A.PathGradientFill() { Path = A.PathShadeValues.Circle },
                            new A.TileRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 }
                        ) { RotateWithShape = true }
                    )
                ) { Name = "Office" }
            );

            theme.Append(themeElements);
            theme.Append(new A.ObjectDefaults());
            theme.Append(new A.ExtraColorSchemeList());

            return theme;
        }

        /// <summary>
        /// Adds shapes from JSON to a slide's shape tree.
        /// Optimized to reduce exception handling overhead and improve performance.
        /// </summary>
        /// <param name="shapeTree">The shape tree to add shapes to</param>
        /// <param name="shapes">List of shape data from JSON</param>
        /// <param name="partContainer">The container part (SlidePart or SlideMasterPart)</param>
        private void AddShapesToSlide(ShapeTree shapeTree, List<ShapeJsonData> shapes, OpenXmlPartContainer partContainer)
        {
            if (shapes == null || shapes.Count == 0)
                return;

            uint shapeIdCounter = GetNextShapeId(shapeTree);
            
            // Pre-determine part type to avoid repeated type checks
            var slidePart = partContainer as SlidePart;
            var isMasterPart = partContainer is SlideMasterPart;

            foreach (var shapeData in shapes)
            {
                try
                {
                    OpenXmlElement shape = null;

                    // Use pre-determined part type for better performance
                    if (slidePart != null)
                    {
                        shape = _slideWriter.CreateShapeFromJson(shapeData, shapeIdCounter++, slidePart);
                    }
                    else if (isMasterPart)
                    {
                        // For master parts, create shapes without image support
                        shape = CreateShapeForMaster(shapeData, shapeIdCounter++);
                    }

                    if (shape != null)
                    {
                        shapeTree.Append(shape);
                    }
                }
                catch (ConversionException ex)
                {
                    // Re-throw conversion exceptions with additional context
                    throw new ConversionException(
                        $"Failed to create shape '{shapeData.Name}': {ex.Message}",
                        $"Shape Creation - {ex.Context}",
                        ex
                    );
                }
                catch (Exception ex)
                {
                    throw new ConversionException(
                        $"Failed to create shape '{shapeData.Name}': {ex.Message}",
                        "Shape Creation",
                        ex
                    );
                }
            }
        }

        /// <summary>
        /// Creates a shape for master slides (without slide part dependency)
        /// </summary>
        private OpenXmlElement CreateShapeForMaster(ShapeJsonData shapeData, uint id)
        {
            // For master slides, we create shapes without requiring a SlidePart
            // This is a simplified version that doesn't support pictures
            if (shapeData.Type?.ToLower() == "picture")
            {
                // Skip pictures on master slides for now
                return null;
            }

            // Use a null slide part - the CreateShapeFromJson will handle basic shapes
            return _slideWriter.CreateShapeFromJson(shapeData, id, null);
        }

        /// <summary>
        /// Gets the next available shape ID in a shape tree.
        /// Optimized to reduce repeated iterations over child elements.
        /// </summary>
        /// <param name="shapeTree">The shape tree to search</param>
        /// <returns>The next available shape ID</returns>
        private uint GetNextShapeId(ShapeTree shapeTree)
        {
            uint maxId = 1;
            
            // Use LINQ for more efficient iteration
            var shapeIds = shapeTree.ChildElements.OfType<P.Shape>()
                .Select(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value)
                .Where(id => id.HasValue)
                .Select(id => id.Value);
            
            var pictureIds = shapeTree.ChildElements.OfType<P.Picture>()
                .Select(p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value)
                .Where(id => id.HasValue)
                .Select(id => id.Value);
            
            var frameIds = shapeTree.ChildElements.OfType<P.GraphicFrame>()
                .Select(f => f.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value)
                .Where(id => id.HasValue)
                .Select(id => id.Value);
            
            var connIds = shapeTree.ChildElements.OfType<P.ConnectionShape>()
                .Select(c => c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id?.Value)
                .Where(id => id.HasValue)
                .Select(id => id.Value);
            
            // Combine all IDs and find max
            var allIds = shapeIds.Concat(pictureIds).Concat(frameIds).Concat(connIds);
            if (allIds.Any())
            {
                maxId = allIds.Max();
            }
            
            return maxId + 1;
        }
    }
}
