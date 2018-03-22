-module(xlsx_util).

-export([new/0, write/2, mkdir/2, get_output_stream/2, write/3, get_sheets/1]).

-record(xlsx, {tmp, files=[], sheets=[]}).

%% @doc Create new xlsx doc
new() ->
    {ok, Dir} = mktemp_dir("xlsx"),
    #xlsx{tmp=Dir}.

get_sheets(#xlsx{sheets=Sheets}) ->
    Sheets.

%% @doc Write zip file contents
write(#xlsx{tmp=Tmp, files=Files}, OutFile) ->
    {ok, CurrentDir} = file:get_cwd(),
    file:set_cwd(Tmp),
    {ok, Path} = zip:create(filename:join(CurrentDir, OutFile), Files),
    os:cmd("rm -rf " ++ Tmp),
    file:set_cwd(CurrentDir),
    {ok, Path}.

mkdir(#xlsx{tmp=Tmp}, Dir) ->
    lists:foldl(
      fun(Part, Path) ->
              D = filename:join(Path, Part),
              case filelib:is_dir(D) of
                  true ->
                      nop;
                  false ->
                      file:make_dir(D)
              end,
              D
      end,
      Tmp,
      string:tokens(Dir, "/")).

get_output_stream(X=#xlsx{tmp=Tmp, files=Files}, RelPath) ->
    mkdir(X, filename:dirname(RelPath)),
    Path = filename:join(Tmp, RelPath),
    {ok, F} = file:open(Path, [write]),
    {ok, {F, X#xlsx{files=[RelPath|Files]}}}.

write(X, RelPath, Bytes) ->
    {ok, {F, X2}} = get_output_stream(X, RelPath),
    ok = file:write(F, Bytes),
    ok = file:close(F),
    {ok, X2}.

%% helpers

mktemp_dir(Prefix) ->
    Rand = integer_to_list(binary:decode_unsigned(crypto:strong_rand_bytes(8)), 36),
    TempPath = filename:basedir(user_cache, Prefix ++ Rand),
    ok = filelib:ensure_dir(TempPath),
    Result = file:make_dir(TempPath),
    case Result of
        ok -> {ok, TempPath};
        Error -> Error
    end.
