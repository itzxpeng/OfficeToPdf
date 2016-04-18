/* ==============================================================================
 * Description：PreviewImageInfo  
 * Author：peytonzhang
 * Created：3/28/2016 11:06:12 AM
 * ==============================================================================*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Convertion
{
    /// <summary>
    /// It is an transmission type for communication
    /// </summary>
    [DataContract]
    public class PreviewImageInfo : ISerializable<XElement>
    {
        [DataMember]
        public string UniqeId
        {
            get;
            set;
        }

        [DataMember]
        public int OrignalIndex
        {
            get;
            set;
        }

        [DataMember]
        public string BlobId
        {
            get;
            set;
        }

        [DataMember]
        public bool IsFinalImage
        {
            get;
            set;
        }

        public byte[] ImageDatas { get; set; }

        XElement ISerializable<XElement>.Serialize()
        {
            XElement xe = new XElement("PagesBlob");
            xe.SetAttributeValue("UniqeId", UniqeId);
            xe.SetAttributeValue("PageIndex", OrignalIndex);
            xe.SetAttributeValue("BlobId", BlobId);
            xe.SetAttributeValue("IsFinalImage", IsFinalImage);
            return xe;
        }

        public override bool Equals(object obj)
        {
            var info = obj as PreviewImageInfo;
            if (info != null)
            {
                return info.UniqeId == this.UniqeId;
            }
            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            if (!string.IsNullOrEmpty(UniqeId))
            {
                return UniqeId.GetHashCode();
            }
            return base.GetHashCode();
        }

        void ISerializable<XElement>.Deserialize(XElement serializationdata)
        {
            if (serializationdata != null)
            {
                var attr = serializationdata.Attribute(XName.Get("PageIndex"));
                if (attr != null)
                {
                    OrignalIndex = int.Parse(attr.Value);
                }

                attr = serializationdata.Attribute(XName.Get("BlobId"));
                if (attr != null)
                {
                    BlobId = attr.Value;
                }

                attr = serializationdata.Attribute(XName.Get("UniqeId"));
                if (attr != null)
                {
                    UniqeId = attr.Value;
                }

                attr = serializationdata.Attribute(XName.Get("IsFinalImage"));
                if (attr != null)
                {
                    IsFinalImage = bool.Parse(attr.Value);
                }
            }
        }
    }
}
