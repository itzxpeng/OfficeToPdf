/* ==============================================================================
 * Description：ISerializable  
 * Author：peytonzhang
 * Created：3/28/2016 11:06:48 AM
 * ==============================================================================*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Convertion
{
    public interface ISerializable<T>
    {
        T Serialize();
        void Deserialize(T serializationdata);
    }
}
