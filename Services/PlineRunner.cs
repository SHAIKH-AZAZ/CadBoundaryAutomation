using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace CadBoundaryAutomation.Services
{
    public class PlineRunner
    {
        private readonly Document _doc;
        private readonly Database _db;

        public ObjectId CreatedBoundaryId { get; private set; } = ObjectId.Null;

        private ObjectEventHandler _appendedHandler;
        private CommandEventHandler _endedHandler;
        private CommandEventHandler _cancelledHandler;
        private CommandEventHandler _failedHandler;
        private EventHandler _idleHandler;

        public event Action Completed;
        public event Action Cancelled;

        public PlineRunner(Document doc)
        {
            _doc = doc;
            _db = doc.Database;
        }

        public void Start()
        {
            _appendedHandler = (sender, e) =>
            {
                if (e.DBObject is Entity ent && ent.OwnerId == _db.CurrentSpaceId)
                {
                    if (ent is Polyline || ent is Polyline2d || ent is Polyline3d)
                        CreatedBoundaryId = ent.ObjectId;
                }
            };
            _db.ObjectAppended += _appendedHandler;

            _endedHandler = (s, e) =>
            {
                if (!IsPlineCommand(e)) return;

                DetachCommandHandlers();

                _idleHandler = (ss, ee) =>
                {
                    AcAp.Idle -= _idleHandler;
                    _idleHandler = null;

                    Completed?.Invoke();
                    DetachAll();
                };

                AcAp.Idle += _idleHandler;
            };

            _cancelledHandler = (s, e) =>
            {
                if (!IsPlineCommand(e)) return;
                Cancelled?.Invoke();
                DetachAll();
            };

            _failedHandler = (s, e) =>
            {
                if (!IsPlineCommand(e)) return;
                Cancelled?.Invoke();
                DetachAll();
            };

            _doc.CommandEnded += _endedHandler;
            _doc.CommandCancelled += _cancelledHandler;
            _doc.CommandFailed += _failedHandler;

            _doc.SendStringToExecute("_.PLINE ", true, false, false);
        }

        private bool IsPlineCommand(CommandEventArgs e)
        {
            string cmd = (e.GlobalCommandName ?? "").Trim();
            return cmd.Equals("PLINE", StringComparison.OrdinalIgnoreCase) ||
                   cmd.Equals("_.PLINE", StringComparison.OrdinalIgnoreCase);
        }

        private void DetachCommandHandlers()
        {
            if (_endedHandler != null) _doc.CommandEnded -= _endedHandler;
            if (_cancelledHandler != null) _doc.CommandCancelled -= _cancelledHandler;
            if (_failedHandler != null) _doc.CommandFailed -= _failedHandler;

            _endedHandler = null;
            _cancelledHandler = null;
            _failedHandler = null;
        }

        private void DetachAll()
        {
            if (_idleHandler != null)
            {
                AcAp.Idle -= _idleHandler;
                _idleHandler = null;
            }

            DetachCommandHandlers();

            if (_appendedHandler != null)
            {
                _db.ObjectAppended -= _appendedHandler;
                _appendedHandler = null;
            }
        }
    }
}
