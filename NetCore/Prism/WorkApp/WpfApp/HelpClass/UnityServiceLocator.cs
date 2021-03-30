using CommonServiceLocator;
using System;
using System.Collections.Generic;
using Unity;
using Unity.Lifetime;
using Unity.Resolution;

namespace WpfApp.HelpClass
{
    public sealed class UnityServiceLocator : ServiceLocatorImplBase, IDisposable
    {
        private IUnityContainer container;
        public UnityServiceLocator(IUnityContainer container)
        {
            this.container = container;
            container.RegisterInstance<IServiceLocator>(this, new ExternallyControlledLifetimeManager());
        }
        public void Dispose()
        {
            if (this.container != null)
            {
                this.container.Dispose();
                this.container = null;
            }
        }

        protected override IEnumerable<object> DoGetAllInstances(Type serviceType)
        {
            if (this.container == null)
            {
                throw new ObjectDisposedException("container");
            }
            return this.container.ResolveAll(serviceType, new ResolverOverride[0]);
        }

        protected override object DoGetInstance(Type serviceType, string key)
        {
            if (this.container == null)
            {
                throw new ObjectDisposedException("container");
            }
            return this.container.Resolve(serviceType, key, new ResolverOverride[0]);
        }

    }
}
