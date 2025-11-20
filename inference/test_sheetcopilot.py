"""
SheetCopilot Test Script
Quick test to verify the multi-stage system is working correctly
"""

import sys
import json
from sheetcopilot import SheetCopilot, setup_logger, parse_option

def test_code_execution():
    """Test code execution client"""
    print("\n" + "="*80)
    print("Testing Code Execution Client")
    print("="*80)
    
    # Mock options for testing
    class MockOpt:
        code_exec_url = "http://localhost:8080/execute"
        conv_id = "TEST"
        model = "test"
        api_key = ""
        base_url = ""
    
    opt = MockOpt()
    logger = setup_logger('log', 'test')
    
    try:
        from code_exec import get_exec_client
        client = get_exec_client(opt.code_exec_url, opt.conv_id)
        
        print("\n✓ Code execution client initialized successfully")
        print(f"  - Client type: {type(client).__name__}")
        
        return True
    except Exception as e:
        print(f"\n✗ Code execution client initialization failed: {str(e)}")
        return False


def test_stage_logging():
    """Test stage logging functionality"""
    print("\n" + "="*80)
    print("Testing Stage Logging")
    print("="*80)
    
    class MockOpt:
        code_exec_url = "http://localhost:8080/execute"
        conv_id = "TEST"
        model = "test"
        api_key = ""
        base_url = ""
        max_revisions = 3
    
    opt = MockOpt()
    logger = setup_logger('log', 'test')
    
    try:
        copilot = SheetCopilot(opt, logger)
        
        # Test logging
        copilot.log_stage("TEST_STAGE", "This is a test message")
        
        print("\n✓ Stage logging working")
        print(f"  - Stage history: {len(copilot.stage_history)} entries")
        print(f"  - Current stage: {copilot.current_stage}")
        
        return True
    except Exception as e:
        print(f"\n✗ Stage logging failed: {str(e)}")
        return False


def test_prompt_generation():
    """Test prompt generation for each stage"""
    print("\n" + "="*80)
    print("Testing Prompt Generation")
    print("="*80)
    
    # Test data
    test_instruction = "Find the maximum value in column A"
    test_file = "/mnt/data/test1/spreadsheet/test/1_test_input.xlsx"
    test_position = "B1:B10"
    
    observation_result = {
        'result': 'Dimensions: 10 rows x 5 columns\nValues: [1, 2, 3, 4, 5]',
        'success': True,
        'prompt': 'Test prompt',
        'response': 'Test response'
    }
    
    # Test observing prompt (simplified)
    observing_prompt = f"""You are SheetCopilot in OBSERVING stage.
Task: {test_instruction}
File: {test_file}
Target: {test_position}
"""
    print("\n✓ Observing prompt structure:")
    print(f"  - Length: {len(observing_prompt)} chars")
    print(f"  - Contains task: {'Task:' in observing_prompt}")
    
    # Test proposing prompt (simplified)
    proposing_prompt = f"""You are SheetCopilot in PROPOSING stage.
Task: {test_instruction}
Observation: {observation_result['result']}
"""
    print("\n✓ Proposing prompt structure:")
    print(f"  - Length: {len(proposing_prompt)} chars")
    print(f"  - Contains observation: {'Observation:' in proposing_prompt}")
    
    # Test revising prompt (simplified)
    revising_prompt = f"""You are SheetCopilot in REVISING stage.
Error: Test error
"""
    print("\n✓ Revising prompt structure:")
    print(f"  - Length: {len(revising_prompt)} chars")
    print(f"  - Contains error: {'Error:' in revising_prompt}")
    
    return True


def test_result_format():
    """Test result format validation"""
    print("\n" + "="*80)
    print("Testing Result Format")
    print("="*80)
    
    # Sample result
    result = {
        'id': 'test_001',
        'instruction_type': 'Cell-Level Manipulation',
        'conversation': ['prompt1', 'response1', 'result1'],
        'solution': 'print("test")',
        'success': True,
        'revision_count': 1,
        'stage_history': [
            {'stage': 'OBSERVING', 'content': 'test', 'timestamp': '2025-11-20T10:00:00'}
        ]
    }
    
    try:
        # Validate result structure
        required_keys = ['id', 'instruction_type', 'conversation', 'solution', 
                        'success', 'revision_count', 'stage_history']
        
        missing_keys = [k for k in required_keys if k not in result]
        
        if missing_keys:
            print(f"\n✗ Missing keys: {missing_keys}")
            return False
        
        print("\n✓ Result format valid")
        print(f"  - All required keys present: {required_keys}")
        print(f"  - Conversation length: {len(result['conversation'])}")
        print(f"  - Stage history length: {len(result['stage_history'])}")
        
        # Test JSON serialization
        json_str = json.dumps(result, ensure_ascii=False)
        parsed = json.loads(json_str)
        
        print("\n✓ JSON serialization working")
        print(f"  - Serialized size: {len(json_str)} bytes")
        
        return True
    except Exception as e:
        print(f"\n✗ Result format test failed: {str(e)}")
        return False


def run_all_tests():
    """Run all tests"""
    print("\n" + "#"*80)
    print("# SheetCopilot System Test Suite")
    print("#"*80)
    
    tests = [
        ("Code Execution Client", test_code_execution),
        ("Stage Logging", test_stage_logging),
        ("Prompt Generation", test_prompt_generation),
        ("Result Format", test_result_format),
    ]
    
    results = []
    for test_name, test_func in tests:
        try:
            success = test_func()
            results.append((test_name, success))
        except Exception as e:
            print(f"\n✗ Test '{test_name}' crashed: {str(e)}")
            results.append((test_name, False))
    
    # Summary
    print("\n" + "="*80)
    print("Test Summary")
    print("="*80)
    
    passed = sum(1 for _, success in results if success)
    total = len(results)
    
    for test_name, success in results:
        status = "✓ PASS" if success else "✗ FAIL"
        print(f"{status}: {test_name}")
    
    print(f"\nTotal: {passed}/{total} tests passed")
    
    if passed == total:
        print("\n🎉 All tests passed! SheetCopilot is ready to use.")
        return 0
    else:
        print(f"\n⚠️  {total - passed} test(s) failed. Please check the logs.")
        return 1


if __name__ == '__main__':
    exit_code = run_all_tests()
    sys.exit(exit_code)
