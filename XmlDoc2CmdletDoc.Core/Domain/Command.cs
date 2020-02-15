using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Reflection;

namespace XmlDoc2CmdletDoc.Core.Domain
{
    /// <summary>
    /// Represents a single cmdlet.
    /// </summary>
    public class Command
    {
        private readonly CmdletAttribute _attribute;
        public readonly Type _cmdletType;
        private readonly Lazy<Type[]> _outputTypes;
        private readonly Lazy<Parameter[]> _parameters;

        /// <summary>
        /// Creates a new instance based on the specified cmdlet type.
        /// </summary>
        /// <param name="cmdletType">The type of the cmdlet. Must be a sub-class of <see cref="Cmdlet"/>
        /// and have a <see cref="CmdletAttribute"/>.</param>
        public Command(Type cmdletType)
        {
            _cmdletType = cmdletType ?? throw new ArgumentNullException(nameof(cmdletType));
            _attribute = _cmdletType.GetCustomAttribute<CmdletAttribute>() ?? throw new ArgumentException("Missing CmdletAttribute", nameof(cmdletType));
            _outputTypes = new Lazy<Type[]>(GetCmdLetTypes);
            _parameters = new Lazy<Parameter[]>(GetParameters);
        }

        private Type[] GetCmdLetTypes()
        {
            var ret = new List<Type>();

            foreach(var attr in Attribute.GetCustomAttributes(_cmdletType, typeof(OutputTypeAttribute)).OfType<OutputTypeAttribute>())
            {
                if (attr.Type == null || attr.Type.Length < 1)
                    throw new Exception($"Type not set for attribute [OutputTypeAttribute] for command {_cmdletType.Name}");

                foreach(var psType in attr.Type)
                {
                    if (psType.Type == null)
                        throw new Exception($"Could not find type for PS type {psType}");

                    ret.Add(psType.Type);
                }
            }

            return ret.Distinct()
                      .OrderBy(type => type.FullName)
                      .ToArray();
        }

        private Parameter[] GetParameters()
        {
            var parameters = _cmdletType.GetMembers(BindingFlags.Instance | BindingFlags.Public)
                                        .Where(member => member.GetCustomAttributes<ParameterAttribute>().Any())
                                        .Select(member => new ReflectionParameter(CmdletType, member))
                                        .ToList<Parameter>();

            if (typeof(IDynamicParameters).IsAssignableFrom(CmdletType))
            {
                foreach (var nestedType in CmdletType.GetNestedTypes(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
                {
                    parameters.AddRange(nestedType.GetMembers(BindingFlags.Instance | BindingFlags.Public)
                                                             .Where(member => member.GetCustomAttributes<ParameterAttribute>().Any())
                                                             .Select(member => new ReflectionParameter(nestedType, member)));
                }

                var cmdlet = (IDynamicParameters)Activator.CreateInstance(CmdletType);

                if (cmdlet.GetDynamicParameters() is RuntimeDefinedParameterDictionary runtimeParamDictionary)
                {
                    parameters.AddRange(runtimeParamDictionary.Where(member => member.Value.Attributes.OfType<ParameterAttribute>().Any())
                                                                         .Select(member => new RuntimeParameter(CmdletType, member.Value)));
                }
            }

            return parameters.ToArray();
        }

        /// <summary>
        /// The type of the cmdlet for this command.
        /// </summary>
        public Type CmdletType => _cmdletType;

        /// <summary>
        /// The cmdlet verb.
        /// </summary>
        public string Verb => _attribute.VerbName;

        /// <summary>
        /// The cmdlet noun.
        /// </summary>
        public string Noun => _attribute.NounName;

        /// <summary>
        /// The cmdlet name, of the form verb-noun.
        /// </summary>
        public string Name => Verb + "-" + Noun;

        /// <summary>
        /// The output types declared by the command.
        /// </summary>
        public Type[] OutputTypes => _outputTypes.Value;

        /// <summary>
        /// The parameters belonging to the command.
        /// </summary>
        public Parameter[] Parameters => _parameters.Value;

        /// <summary>
        /// The command's parameters that belong to the specified parameter set.
        /// </summary>
        /// <param name="parameterSetName">The name of the parameter set.</param>
        /// <returns>
        /// The command's parameters that belong to the specified parameter set.
        /// </returns>
        public IEnumerable<Parameter> GetParameters(string parameterSetName) =>
            parameterSetName == ParameterAttribute.AllParameterSets
                ? Parameters
                : Parameters.Where(p => p.ParameterSetNames.Contains(parameterSetName) ||
                                        p.ParameterSetNames.Contains(ParameterAttribute.AllParameterSets));

        /// <summary>
        /// The names of the parameter sets that the parameters belongs to.
        /// </summary>
        public IEnumerable<string> ParameterSetNames =>
            Parameters.SelectMany(p => p.ParameterSetNames)
                      .Union(new[] {ParameterAttribute.AllParameterSets}) // Parameterless cmdlets need this seeded
                      .Distinct();
    }
}
