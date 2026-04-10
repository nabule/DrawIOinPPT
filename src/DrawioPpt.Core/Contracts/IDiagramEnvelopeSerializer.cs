using DrawioPpt.Core.Models;

namespace DrawioPpt.Core.Contracts
{
    public interface IDiagramEnvelopeSerializer
    {
        string Serialize(DiagramEnvelope envelope);
        DiagramEnvelope Deserialize(string content);
        bool CanDeserialize(string content);
    }
}

