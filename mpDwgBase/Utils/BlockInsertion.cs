namespace mpDwgBase.Utils
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Autodesk.AutoCAD.ApplicationServices.Core;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.Internal;
    using Models;
    using ModPlusAPI;

    public static class BlockInsertion
    {
        private const string LangItem = "mpDwgBase";

        public static Dictionary<AttributeDefinition, AttributeReference> AppendAttribToBlock(
            Transaction tr, BlockReference blkref, List<string> atts)
        {
            var blockTableRecord = (BlockTableRecord)tr.GetObject(blkref.BlockTableRecord, OpenMode.ForRead);
            var i = 0;
            if (blockTableRecord.HasAttributeDefinitions)
            {
                var attributeReferences = new Dictionary<AttributeDefinition, AttributeReference>();
                var attributeDefinitions = blockTableRecord.Cast<ObjectId>()
                    .Select(id => tr.GetObject(id, OpenMode.ForRead))
                    .OfType<AttributeDefinition>()
                    .Where(ad => !(ad.Constant || ad.Invisible));
                foreach (var attributeDefinition in attributeDefinitions)
                {
                    var attributeReference = new AttributeReference();
                    attributeReference.SetAttributeFromBlock(attributeDefinition, blkref.BlockTransform);
                    if (i < atts.Count)
                    {
                        attributeReference.TextString = atts[i];
                    }
                    else
                    {
                        attributeReference.TextString = attributeDefinition.TextString;
                    }

                    i++;
                    blkref.AttributeCollection.AppendAttribute(attributeReference);
                    tr.AddNewlyCreatedDBObject(attributeReference, true);
                    attributeReferences.Add(attributeDefinition, attributeReference);
                }

                return attributeReferences;
            }

            return null;
        }

        /// <summary>
        /// Вставка блока с атрибутами
        /// </summary>
        /// <param name="promptCounter">0 - только вставка, 1 - с поворотом</param>
        /// <param name="tr">Транзакция</param>
        /// <param name="db">База данных чертежа</param>
        /// <param name="ed">Editor</param>
        /// <param name="blkDefId">ObjectId блока</param>
        /// <param name="attributes">Список имен атрибутов</param>
        /// <param name="isAnnotative">Аннотативность блока-исходника</param>
        /// <returns></returns>
        public static ObjectId InsertBlockRef(
            int promptCounter,
            Transaction tr,
            Database db,
            Editor ed,
            ObjectId blkDefId,
            List<string> attributes,
            bool isAnnotative)
        {
            var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            var blockReference = new BlockReference(Point3d.Origin, blkDefId);
            if (isAnnotative)
            {
                var ocm = db.ObjectContextManager;
                var occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");
                ObjectContexts.AddContext(blockReference, occ.CurrentContext);
            }

            var id = btr.AppendEntity(blockReference);
            tr.AddNewlyCreatedDBObject(blockReference, true);
            var jig = new BlockRefJig(blockReference, AppendAttribToBlock(tr, blockReference, attributes));
            jig.SetPromptCounter(0);
            var res = ed.Drag(jig);
            if (res.Status == PromptStatus.OK)
            {
                if (promptCounter == 1)
                {
                    jig.SetPromptCounter(promptCounter);
                    res = ed.Drag(jig);
                    if (res.Status == PromptStatus.OK)
                    {
                        return id;
                    }
                }
                else
                {
                    return id;
                }
            }

            blockReference.Erase();
            return ObjectId.Null;
        }

        internal class BlockRefJig : EntityJig
        {
            private Point3d _position;
            private Point3d _basePoint;
            private double _angle;
            private int _promptCounter;
            private Matrix3d _ucs;
            private Matrix3d _matrix3d;
            private Editor _ed = Application.DocumentManager.MdiActiveDocument.Editor;
            private Dictionary<AttributeDefinition, AttributeReference> _attributeReferences;

            public BlockRefJig(
                BlockReference blkReference,
                Dictionary<AttributeDefinition, AttributeReference> attributeReferences)
                : base(blkReference)
            {
                _position = default;
                _angle = 0;
                _ucs = _ed.CurrentUserCoordinateSystem;
                _attributeReferences = attributeReferences;
                Update();
            }

            /// <inheritdoc />
            protected override SamplerStatus Sampler(JigPrompts prompts)
            {
                switch (_promptCounter)
                {
                    case 0:
                    {
                        var jigOpts = new JigPromptPointOptions($"\n{Language.GetItem(LangItem, "msg13")}:");
                        jigOpts.UserInputControls =
                            UserInputControls.Accept3dCoordinates |
                            UserInputControls.NoZeroResponseAccepted |
                            UserInputControls.NoNegativeResponseAccepted;
                        var res = prompts.AcquirePoint(jigOpts);
                        var pnt = res.Value;
                        if (pnt != _position)
                        {
                            _position = pnt;
                            _basePoint = _position;
                        }
                        else
                        {
                            return SamplerStatus.NoChange;
                        }

                        if (res.Status == PromptStatus.Cancel)
                        {
                            return SamplerStatus.Cancel;
                        }

                        return SamplerStatus.OK;
                    }

                    case 1:
                    {
                        var jigOpts = new JigPromptAngleOptions($"\n{Language.GetItem(LangItem, "msg14")}:");
                        jigOpts.UserInputControls =
                            UserInputControls.Accept3dCoordinates |
                            UserInputControls.NoNegativeResponseAccepted |
                            UserInputControls.GovernedByUCSDetect |
                            UserInputControls.UseBasePointElevation;
                        jigOpts.Cursor = CursorType.RubberBand;
                        jigOpts.UseBasePoint = true;
                        jigOpts.BasePoint = _basePoint;
                        var res = prompts.AcquireAngle(jigOpts);
                        var angleTemp = res.Value;
                        if (Math.Abs(angleTemp - _angle) > Tolerance.Global.EqualVector)
                        {
                            _angle = angleTemp;
                        }
                        else
                        {
                            return SamplerStatus.NoChange;
                        }

                        if (res.Status == PromptStatus.Cancel)
                        {
                            return SamplerStatus.Cancel;
                        }

                        return SamplerStatus.OK;
                    }

                    default:
                        return SamplerStatus.NoChange;
                }
            }

            /// <inheritdoc />
            protected sealed override bool Update()
            {
                try
                {
                    var blockReference = (BlockReference)Entity;
                    blockReference.Normal = Vector3d.ZAxis;
                    blockReference.Position = _position.TransformBy(_ed.CurrentUserCoordinateSystem.Inverse());
                    blockReference.Rotation = _angle;
                    blockReference.TransformBy(_ucs);
                    if (_attributeReferences != null)
                    {
                        _matrix3d = blockReference.BlockTransform;
                        foreach (var att in _attributeReferences)
                        {
                            var attributeReference = att.Value;
                            var s = attributeReference.TextString;
                            attributeReference.SetAttributeFromBlock(att.Key, _matrix3d);
                            attributeReference.TextString = s;
                        }
                    }
                }
                catch
                {
                    return false;
                }

                return true;
            }

            public void SetPromptCounter(int i)
            {
                if (i == 0 || i == 1)
                {
                    _promptCounter = i;
                }
            }
        }

        /// <summary>
        /// Добавление атрибутов для спецификации
        /// </summary>
        public static void AddAttributesForSpecification(
            Transaction tr, BlockReference blkRef, DwgBaseItem dwgBaseItem)
        {
            if (HasAttributesForSpecification(tr, blkRef.ObjectId))
            {
                // update attributes
                foreach (ObjectId id in blkRef.AttributeCollection)
                {
                    var attr = tr.GetObject(id, OpenMode.ForWrite) as AttributeReference;
                    if (attr != null)
                    {
                        if (attr.Tag.ToLower().Equals("mp:позиция") ||
                            attr.Tag.ToLower().Equals("mp:position"))
                        {
                            attr.TextString = dwgBaseItem.PositionValue;
                        }

                        if (attr.Tag.ToLower().Equals("mp:обозначение") ||
                            attr.Tag.ToLower().Equals("mp:designation"))
                        {
                            attr.TextString = dwgBaseItem.DesignationValue;
                        }

                        if (attr.Tag.ToLower().Equals("mp:наименование") ||
                            attr.Tag.ToLower().Equals("mp:name"))
                        {
                            attr.TextString = dwgBaseItem.NominationValue;
                        }

                        if (attr.Tag.ToLower().Equals("mp:масса") ||
                            attr.Tag.ToLower().Equals("mp:mass"))
                        {
                            attr.TextString = dwgBaseItem.MassValue;
                        }

                        if (attr.Tag.ToLower().Equals("mp:примечание") ||
                            attr.Tag.ToLower().Equals("mp:note"))
                        {
                            attr.TextString = dwgBaseItem.NoteValue;
                        }
                    }
                }
            }
            else
            {
                // add attributes
                BlockTableRecord acBlkTblRec;
                if (blkRef.IsDynamicBlock) // get the real dynamic block name.
                {
                    acBlkTblRec = tr.GetObject(blkRef.DynamicBlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                }
                else
                {
                    acBlkTblRec = tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                }

                if (acBlkTblRec != null)
                {
                    // blkName = acBlkTblRec.Name;
                    var addAttrDefenitions = AddAttrDefinitions(acBlkTblRec, blkRef, tr);
                    if (addAttrDefenitions.Any())
                    {
                        foreach (ObjectId id in blkRef.AttributeCollection)
                        {
                            var attr = tr.GetObject(id, OpenMode.ForWrite) as AttributeReference;
                            if (attr != null)
                            {
                                if (attr.Tag.ToLower().Equals("mp:позиция") ||
                                    attr.Tag.ToLower().Equals("mp:position"))
                                {
                                    attr.TextString = dwgBaseItem.PositionValue;
                                }

                                if (attr.Tag.ToLower().Equals("mp:обозначение") ||
                                    attr.Tag.ToLower().Equals("mp:designation"))
                                {
                                    attr.TextString = dwgBaseItem.DesignationValue;
                                }

                                if (attr.Tag.ToLower().Equals("mp:наименование") ||
                                    attr.Tag.ToLower().Equals("mp:name"))
                                {
                                    attr.TextString = dwgBaseItem.NominationValue;
                                }

                                if (attr.Tag.ToLower().Equals("mp:масса") ||
                                    attr.Tag.ToLower().Equals("mp:mass"))
                                {
                                    attr.TextString = dwgBaseItem.MassValue;
                                }

                                if (attr.Tag.ToLower().Equals("mp:примечание") ||
                                    attr.Tag.ToLower().Equals("mp:note"))
                                {
                                    attr.TextString = dwgBaseItem.NoteValue;
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Проверка, что блок имеет атрибуты для заполнения спецификации
        /// </summary>
        private static bool HasAttributesForSpecification(Transaction tr, ObjectId objectId)
        {
            var allowAttributesTags = new List<string> { "mp:position", "mp:designation", "mp:name", "mp:mass", "mp:note" };
            if (Language.RusWebLanguages.Contains(Language.CurrentLanguageName))
            {
                allowAttributesTags = new List<string> { "mp:позиция", "mp:обозначение", "mp:наименование", "mp:масса", "mp:примечание" };
            }

            var blk = tr.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            if (blk != null)
            {
                // Если это блок
                if (blk.AttributeCollection.Count > 0)
                {
                    foreach (ObjectId id in blk.AttributeCollection)
                    {
                        var attr = tr.GetObject(id, OpenMode.ForRead) as AttributeReference;
                        if (allowAttributesTags.Contains(attr?.Tag.ToLower()))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        // Add an attribute definition to the block
        private static List<AttributeDefinition> AddAttrDefinitions(BlockTableRecord acBlkTblRec, BlockReference blkRef, Transaction tr)
        {
            var attributeDefinitions = new List<AttributeDefinition>();
            var allowAttributesTags = new List<string> { "mp:position", "mp:designation", "mp:name", "mp:mass", "mp:note" };
            if (Language.RusWebLanguages.Contains(Language.CurrentLanguageName))
            {
                allowAttributesTags = new List<string> { "mp:позиция", "mp:обозначение", "mp:наименование", "mp:масса", "mp:примечание" };
            }

            var allowAttributesPrompt = new List<string>
            {
                Language.GetItem(LangItem, "ad1"), Language.GetItem(LangItem, "ad2"),
                Language.GetItem(LangItem, "ad3"), Language.GetItem(LangItem, "ad4"), Language.GetItem(LangItem, "ad5")
            };

            for (var i = 0; i < allowAttributesTags.Count; i++)
            {
                using (var acAttDef = new AttributeDefinition())
                {
                    acAttDef.Position = new Point3d(0, 0, 0);
                    acAttDef.Invisible = true;
                    acAttDef.Prompt = allowAttributesPrompt[i];
                    acAttDef.Tag = allowAttributesTags[i];
                    acAttDef.TextString = string.Empty;
                    acBlkTblRec.AppendEntity(acAttDef);

                    var attRef = new AttributeReference();
                    attRef.SetAttributeFromBlock(acAttDef, blkRef.BlockTransform);
                    blkRef.AttributeCollection.AppendAttribute(attRef);
                    tr.AddNewlyCreatedDBObject(attRef, true);

                    attributeDefinitions.Add(acAttDef);
                }
            }

            return attributeDefinitions;
        }
    }
}