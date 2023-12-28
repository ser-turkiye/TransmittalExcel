package junit

import de.ser.doxis4.agentserver.AgentExecutionResult
import org.junit.*
import ser.TransmittalLoad

class TEST_TransmittalLoad {

    Binding binding

    @BeforeClass
    static void initSessionPool() {
        AgentTester.initSessionPool()
    }

    @Before
    void retrieveBinding() {
        binding = AgentTester.retrieveBinding()
    }

    @Test
    void testForAgentResult() {
        def agent = new TransmittalLoad();

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "ST03BPM2428438103-3b29-48bb-ba1e-01745d9202b5182023-12-28T06:58:30.480Z010"

        def result = (AgentExecutionResult) agent.execute(binding.variables)
        assert result.resultCode == 0
    }

    @Test
    void testForJavaAgentMethod() {
        //def agent = new JavaAgent()
        //agent.initializeGroovyBlueline(binding.variables)
        //assert agent.getServerVersion().contains("Linux")
    }

    @After
    void releaseBinding() {
        AgentTester.releaseBinding(binding)
    }

    @AfterClass
    static void closeSessionPool() {
        AgentTester.closeSessionPool()
    }
}
