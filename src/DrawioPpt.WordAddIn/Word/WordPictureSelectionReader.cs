using DrawioPpt.Core.Contracts;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using WordInterop = Microsoft.Office.Interop.Word;

namespace DrawioPpt.WordAddIn.Word
{
    public class WordPictureSelectionReader
    {
        private readonly IDiagramEnvelopeSerializer _serializer;
        private readonly SelectedPictureAccessor _accessor;

        public WordPictureSelectionReader()
            : this(new DiagramEnvelopeSerializer(), new SelectedPictureAccessor())
        {
        }

        public WordPictureSelectionReader(IDiagramEnvelopeSerializer serializer, SelectedPictureAccessor accessor)
        {
            _serializer = serializer;
            _accessor = accessor;
        }

        public SelectionContext Read(WordInterop.Selection selection)
        {
            SelectionContext context = new SelectionContext();
            if (selection == null)
            {
                return context;
            }

            context.HasSelection = true;
            WordPictureReference picture = _accessor.GetSingleSelectedPicture(selection);
            if (picture == null)
            {
                return context;
            }

            context.HasSinglePicture = true;
            context.PictureName = picture.Name;
            context.ReferenceKey = picture.ReferenceKey;
            context.AlternativeText = picture.AlternativeText ?? string.Empty;

            DiagramEnvelope envelope;
            if (_serializer.CanDeserialize(context.AlternativeText))
            {
                envelope = _serializer.Deserialize(context.AlternativeText);
                context.DiagramId = envelope == null ? string.Empty : envelope.DiagramId;
                context.IsManagedPicture = !string.IsNullOrWhiteSpace(context.DiagramId);
            }

            return context;
        }
    }
}
