
class Log(object):
    def write_log(file_name, exception, to_console):
        try:
            error_file = open(file_name, 'w+')
            error_file.write('{}'.format(e) )
            error_file.close()

            if to_console: #prints to screen
                print('{}'.format(e.with_traceback) )

        except Exception as e:
            print('Log: write_error_log(): {}'.format(e.with_traceback ) )

        return None

    def pause():
        input_clear = input()

        return None
