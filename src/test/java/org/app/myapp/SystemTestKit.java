package org.app.myapp;

import no.ks.eventstore2.eventstore.CompleteSubscriptionRegistered;
import akka.actor.ActorSystem;
import akka.testkit.TestActorRef;
import akka.testkit.TestKit;

import org.app.myapp.monitoring.ApplicationStatusProjection;
import org.app.myapp.util.IdUtil;

public class SystemTestKit extends TestKit {

    protected static final ActorSystem actorSystem = ActorSystem.create("testSystem");

    public SystemTestKit() {
        super(actorSystem);
    }

    protected TestActorRef<ApplicationStatusProjection> createApplicationStatusProjection() {
        TestActorRef<ApplicationStatusProjection> projection = TestActorRef.create(actorSystem, ApplicationStatusProjection.mkProps(testActor()), IdUtil.createUUID());
        projection.tell(new CompleteSubscriptionRegistered(null), testActor());
        return projection;
    }
}
