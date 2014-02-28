using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using SpaceClaim.Api.V11;
using SpaceClaim.Api.V11.Display;
using ThreeJsExport.Properties;
using SpaceClaim.Api.V11.Extensibility;
using SpaceClaim.Api.V11.Geometry;
using SpaceClaim.Api.V11.Modeler;
using Point = SpaceClaim.Api.V11.Geometry.Point;
using ScreenPoint = System.Drawing.Point;
using System.Text;
using System.IO;
using Newtonsoft.Json;

namespace SpaceClaim.AddIn.ThreeJsExport {
    class ExportThreeJsToolCapsule : CommandCapsule {
        public ExportThreeJsToolCapsule(string commandName)
            : base(commandName, "Export") {
        }

        protected override void OnInitialize(Command command) {
        }

        protected override void OnUpdate(Command command) {
            Window window = Window.ActiveWindow;
            command.IsEnabled = window != null &&
                                window.Scene is Part;
            command.IsChecked = false;
        }

        public static IEnumerable<IPart> WalkParts(Part part) { // Copied from SpaceClaim.Api.V10.Examples class ShowBomCapsule
            Debug.Assert(part != null);

            // GetDescendants goes not include the object itself
            yield return part;

            foreach (IPart descendant in part.GetDescendants<IPart>())
                yield return descendant;
        }

        protected override void OnExecute(Command command, ExecutionContext context, Rectangle buttonRect) {
            var window = Window.ActiveWindow;
            var document = window.Document;

            var surfaceDeviationCommand = Command.GetCommand("SurfaceDeviation");
            double surfaceDeviation;
            if (!double.TryParse(surfaceDeviationCommand.Text, out surfaceDeviation))
                surfaceDeviation = (new TessellationOptions()).SurfaceDeviation;

            var angleDeviationCommand = Command.GetCommand("AngleDeviation");
            double angleDeviation;
            if (!double.TryParse(angleDeviationCommand.Text, out angleDeviation))
                angleDeviation = (new TessellationOptions()).AngleDeviation;

            Part mainPart = (Part)Window.ActiveWindow.Scene;

            var tessellations = new Dictionary<Body, PartTessellation>();

            var totalTessellation = new PartTessellation();
            foreach (var iPart in WalkParts(mainPart)) {
                var transform = iPart.TransformToMaster.Inverse;

                foreach (var body in iPart.Bodies) {
                    if (!body.GetVisibility(null) ?? !body.Master.Layer.IsVisible(null))
                        continue;

                    var masterBody = body.Master.Shape;

                    PartTessellation tessellation;
                    if (!tessellations.TryGetValue(masterBody, out tessellation)) {
                        Func<Face, Color> getColor = f => body.Master.GetDesignFace(f).GetColor(null) ??
                                            body.Master.GetColor(null) ??
                                            body.Master.Layer.GetColor(null);
                        tessellation = GetBodyTessellation(masterBody, getColor, surfaceDeviation, angleDeviation);
                        tessellations[masterBody] = tessellation;
                    }

                    totalTessellation.Add(tessellation.GetTransformed(transform));
                }
            }

            var result = totalTessellation.ToJson();

            var d = new SaveFileDialog();
            d.FileName = document.Path;
            d.ShowDialog();

            File.WriteAllText(d.FileName, result);
        }

        class PartTessellation {
            public List<BodyTessellation> Lines = new List<BodyTessellation>();
            public List<BodyTessellation> Meshes = new List<BodyTessellation>();

            public PartTessellation GetTransformed(Matrix transform) {
                return new PartTessellation {
                    Lines = new List<BodyTessellation>(this.Lines.Select(l => l.GetTransformed(transform))),
                    Meshes = new List<BodyTessellation>(this.Meshes.Select(l => l.GetTransformed(transform)))
                };
            }

            public void Add(PartTessellation other) {
                Lines.AddRange(other.Lines);
                Meshes.AddRange(other.Meshes);
            }

            public string ToJson() {
                StringBuilder sb = new StringBuilder();
                StringWriter sw = new StringWriter(sb);

                using (JsonWriter writer = new JsonTextWriter(sw)) {
                    writer.Formatting = Formatting.Indented;

                    writer.WriteStartObject();

                    if (Lines.Count > 0) {
                        writer.WritePropertyName("lines");
                        writer.WriteStartArray();

                        foreach (var line in Lines) {
                            writer.WriteRawValue(line.ToJson());
                        }

                        writer.WriteEndArray();
                    }

                    if (Meshes.Count > 0) {
                        writer.WritePropertyName("meshes");
                        writer.WriteStartArray();

                        foreach (var line in Meshes) {
                            writer.WriteRawValue(line.ToJson());
                        }

                        writer.WriteEndArray();
                    }

                    writer.WriteEndObject();
                }

                return sb.ToString();
            }
        }

        class BodyTessellation {
            public List<Point> VertexPositions = new List<Point>();
            public List<Direction> VertexNormals = new List<Direction>();
            public List<Color> FaceColors = new List<Color>();
            public List<FaceStruct> Faces = new List<FaceStruct>();

            public BodyTessellation GetTransformed(Matrix transform) {
                return new BodyTessellation {
                    VertexPositions = new List<Point>(this.VertexPositions.Select(p => transform * p)),
                    VertexNormals = new List<Direction>(this.VertexNormals.Select(n => transform * n)),
                    FaceColors = this.FaceColors,
                    Faces = this.Faces
                };
            }

            public void Add(BodyTessellation other) {
                int vertexOffset = VertexPositions.Count;
                int faceColorOffset = FaceColors.Count;

                VertexPositions.AddRange(other.VertexPositions);
                VertexNormals.AddRange(other.VertexNormals);
                FaceColors.AddRange(other.FaceColors);
                Faces.AddRange(other.Faces.Select(f => new FaceStruct {
                    Vertex1 = f.Vertex1 + vertexOffset,
                    Vertex2 = f.Vertex2 + vertexOffset,
                    Vertex3 = f.Vertex3 + vertexOffset,
                    Color = f.Color + faceColorOffset
                }));
            }

            public string ToJson() {
                StringBuilder sb = new StringBuilder();
                StringWriter sw = new StringWriter(sb);

                using (JsonWriter writer = new JsonTextWriter(sw)) {
                    writer.Formatting = Formatting.Indented;

                    writer.WriteStartObject();

                    writer.WritePropertyName("metadata");
                    writer.WriteStartObject();
                    writer.WritePropertyName("formatVersion");
                    writer.WriteValue(3);
                    writer.WriteEndObject();

                    writer.WritePropertyName("vertices");
                    writer.WriteStartArray();
                    foreach (var vertex in VertexPositions)
                        writer.WriteRawValue(string.Format("{0:0.0###############},{1:0.0###############},{2:0.0###############}", vertex.X, vertex.Y, vertex.Z));
                    writer.WriteEndArray();

                    writer.WritePropertyName("normals");
                    writer.WriteStartArray();
                    foreach (var normal in VertexNormals)
                        writer.WriteRawValue(string.Format("{0:0.0###############},{1:0.0###############},{2:0.0###############}", normal.X, normal.Y, normal.Z));
                    writer.WriteEndArray();

                    writer.WritePropertyName("colors");
                    writer.WriteStartArray();
                    foreach (var color in FaceColors)
                        writer.WriteValue(((long)color.R << 16) | ((long)color.G << 8) | color.B);
                    writer.WriteEndArray();

                    writer.WritePropertyName("faces");
                    writer.WriteStartArray();
                    foreach (var face in Faces)
                        writer.WriteRawValue(face.ToString());
                    writer.WriteEndArray();

                    writer.WriteEndObject();
                }

                return sb.ToString();
            }
        }

        class FaceStruct {
            public int Vertex1;
            public int Vertex2;
            public int Vertex3;
            public int Color;

            [Flags]
            enum FaceType {
                Triangle = 0,
                Quad = 1,
                Material = 2,
                UV = 4,
                VertexUV = 8,
                Normal = 16,
                VertexNormal = 32,
                Color = 64,
                VertexColor = 128
            };

            public override string ToString() {
                var faceType = FaceType.Triangle | FaceType.VertexNormal | FaceType.Color;

                return string.Format("{0}, {1},{2},{3}, {4},{5},{6}, {7}", (int)faceType, Vertex1, Vertex2, Vertex3, Vertex1, Vertex2, Vertex3, Color);
            }
        }

        static PartTessellation GetBodyTessellation(Body body, Func<Face, Color> faceColor, double surfaceDeviation, double angleDeviation) {
            var tessellationOptions = new TessellationOptions(surfaceDeviation, angleDeviation);
            var tessellation = body.GetTessellation(null, tessellationOptions);

            var vertices = new Dictionary<PositionNormalTextured, int>();
            var vertexList = new List<Point>();
            var normalList = new List<Direction>();

            var colors = new Dictionary<Color, int>();
            var colorList = new List<Color>();

            var faces = new List<FaceStruct>();
            foreach (var pair in tessellation) {
                var color = faceColor(pair.Key);

                int colorIndex;
                if (!colors.TryGetValue(color, out colorIndex)) {
                    colorList.Add(color);
                    colorIndex = colorList.Count - 1;
                    colors[color] = colorIndex;
                }

                var vertexIndices = new Dictionary<int, int>();

                var i = 0;
                foreach (var vertex in pair.Value.Vertices) {
                    int index;
                    if (!vertices.TryGetValue(vertex, out index)) {
                        vertexList.Add(vertex.Position);
                        normalList.Add(vertex.Normal);
                        index = vertexList.Count - 1;
                        vertices[vertex] = index;
                    }

                    vertexIndices[i] = index;
                    i++;
                }

                foreach (var facet in pair.Value.Facets) {
                    faces.Add(new FaceStruct {
                        Vertex1 = vertexIndices[facet.Vertex0],
                        Vertex2 = vertexIndices[facet.Vertex1],
                        Vertex3 = vertexIndices[facet.Vertex2],
                        Color = colorIndex
                    });
                }
            }

            var faceTessellation = new BodyTessellation {
                VertexPositions = vertexList,
                VertexNormals = normalList,
                FaceColors = colorList,
                Faces = faces
            };

            List<BodyTessellation> edges = new List<BodyTessellation>();
            foreach (var edge in body.Edges) {
                edges.Add(new BodyTessellation {
                    VertexPositions = new List<Point>(edge.GetPolyline())
                });
            }

            return new PartTessellation {
                Lines = edges,
                Meshes = { faceTessellation }
            };
        }
    }
}